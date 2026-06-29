"use strict";
/*
 * Helper for running LibreOffice (soffice) in environments where AF_UNIX
 * sockets may be blocked (e.g., sandboxed VMs). Detects the restriction at
 * runtime and applies an LD_PRELOAD shim if needed. (Port of soffice.py)
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync, execFileSync } = require("child_process");

const _SHIM_SO = path.join(os.tmpdir(), "lo_socket_shim.so");

function getSofficeEnv() {
  const env = Object.assign({}, process.env);
  env.SAL_USE_VCLPLUGIN = "svp";

  if (needsShim()) {
    const shim = ensureShim();
    env.LD_PRELOAD = String(shim);
  }

  return env;
}

function runSoffice(args, options) {
  const env = getSofficeEnv();
  const opts = Object.assign({ env: env }, options || {});
  if (!opts.env) opts.env = env;
  return spawnSync("soffice", args, opts);
}

function needsShim() {
  // Mirror of Python's socket(AF_UNIX) probe: success => no shim needed.
  const script =
    "const net=require('net'),os=require('os'),path=require('path'),fs=require('fs');" +
    "const p=path.join(os.tmpdir(),'lo_shim_test_'+process.pid+'.sock');" +
    "const s=net.createServer();" +
    "s.on('error',()=>process.exit(1));" +
    "s.listen(p,()=>{s.close(()=>{try{fs.unlinkSync(p)}catch(e){}process.exit(0)})});" +
    "setTimeout(()=>process.exit(1),1000);";
  try {
    const res = spawnSync(process.execPath, ["-e", script], { timeout: 5000 });
    return res.status !== 0;
  } catch (e) {
    return true;
  }
}

function ensureShim() {
  if (fs.existsSync(_SHIM_SO)) {
    return _SHIM_SO;
  }

  const src = path.join(os.tmpdir(), "lo_socket_shim.c");
  fs.writeFileSync(src, _SHIM_SOURCE);
  execFileSync("gcc", ["-shared", "-fPIC", "-o", String(_SHIM_SO), String(src), "-ldl"]);
  fs.unlinkSync(src);
  return _SHIM_SO;
}

const _SHIM_SOURCE = String.raw`
#define _GNU_SOURCE
#include <dlfcn.h>
#include <errno.h>
#include <signal.h>
#include <stdio.h>
#include <stdlib.h>
#include <sys/socket.h>
#include <unistd.h>

static int (*real_socket)(int, int, int);
static int (*real_socketpair)(int, int, int, int[2]);
static int (*real_listen)(int, int);
static int (*real_accept)(int, struct sockaddr *, socklen_t *);
static int (*real_close)(int);
static int (*real_read)(int, void *, size_t);

/* Per-FD bookkeeping (FDs >= 1024 are passed through unshimmed). */
static int is_shimmed[1024];
static int peer_of[1024];
static int wake_r[1024];            /* accept() blocks reading this */
static int wake_w[1024];            /* close()  writes to this      */
static int listener_fd = -1;        /* FD that received listen()    */

__attribute__((constructor))
static void init(void) {
    real_socket     = dlsym(RTLD_NEXT, "socket");
    real_socketpair = dlsym(RTLD_NEXT, "socketpair");
    real_listen     = dlsym(RTLD_NEXT, "listen");
    real_accept     = dlsym(RTLD_NEXT, "accept");
    real_close      = dlsym(RTLD_NEXT, "close");
    real_read       = dlsym(RTLD_NEXT, "read");
    for (int i = 0; i < 1024; i++) {
        peer_of[i] = -1;
        wake_r[i]  = -1;
        wake_w[i]  = -1;
    }
}

/* ---- socket ---------------------------------------------------------- */
int socket(int domain, int type, int protocol) {
    if (domain == AF_UNIX) {
        int fd = real_socket(domain, type, protocol);
        if (fd >= 0) return fd;
        /* socket(AF_UNIX) blocked – fall back to socketpair(). */
        int sv[2];
        if (real_socketpair(domain, type, protocol, sv) == 0) {
            if (sv[0] >= 0 && sv[0] < 1024) {
                is_shimmed[sv[0]] = 1;
                peer_of[sv[0]]    = sv[1];
                int wp[2];
                if (pipe(wp) == 0) {
                    wake_r[sv[0]] = wp[0];
                    wake_w[sv[0]] = wp[1];
                }
            }
            return sv[0];
        }
        errno = EPERM;
        return -1;
    }
    return real_socket(domain, type, protocol);
}

/* ---- listen ---------------------------------------------------------- */
int listen(int sockfd, int backlog) {
    if (sockfd >= 0 && sockfd < 1024 && is_shimmed[sockfd]) {
        listener_fd = sockfd;
        return 0;
    }
    return real_listen(sockfd, backlog);
}

/* ---- accept ---------------------------------------------------------- */
int accept(int sockfd, struct sockaddr *addr, socklen_t *addrlen) {
    if (sockfd >= 0 && sockfd < 1024 && is_shimmed[sockfd]) {
        /* Block until close() writes to the wake pipe. */
        if (wake_r[sockfd] >= 0) {
            char buf;
            real_read(wake_r[sockfd], &buf, 1);
        }
        errno = ECONNABORTED;
        return -1;
    }
    return real_accept(sockfd, addr, addrlen);
}

/* ---- close ----------------------------------------------------------- */
int close(int fd) {
    if (fd >= 0 && fd < 1024 && is_shimmed[fd]) {
        int was_listener = (fd == listener_fd);
        is_shimmed[fd] = 0;

        if (wake_w[fd] >= 0) {              /* unblock accept() */
            char c = 0;
            write(wake_w[fd], &c, 1);
            real_close(wake_w[fd]);
            wake_w[fd] = -1;
        }
        if (wake_r[fd] >= 0) { real_close(wake_r[fd]); wake_r[fd]  = -1; }
        if (peer_of[fd] >= 0) { real_close(peer_of[fd]); peer_of[fd] = -1; }

        if (was_listener)
            _exit(0);                        /* conversion done – exit */
    }
    return real_close(fd);
}
`;

if (require.main === module) {
  const result = runSoffice(process.argv.slice(2), { stdio: "inherit" });
  process.exit(result.status == null ? 1 : result.status);
}

module.exports = { getSofficeEnv, runSoffice };
