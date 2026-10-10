#!/usr/bin/env python3
"""Minimal local forward proxy: Chromium -> 127.0.0.1:8888 (no auth)
   -> upstream authenticated egress proxy (auth handled here in Python).

Run: python forward_proxy.py  (blocking; run in background)
"""
import asyncio, base64, os
from urllib.parse import urlparse, unquote

LISTEN = ("127.0.0.1", 8888)
BUF = 65536


def _upstream():
    v = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
    p = urlparse(v)
    auth = ""
    if p.username:
        cred = f"{unquote(p.username)}:{unquote(p.password or '')}"
        auth = "Proxy-Authorization: Basic " + base64.b64encode(cred.encode()).decode() + "\r\n"
    return p.hostname, p.port, auth


UP_HOST, UP_PORT, UP_AUTH = _upstream()


async def relay(reader, writer):
    try:
        while True:
            data = await reader.read(BUF)
            if not data:
                break
            writer.write(data)
            await writer.drain()
    except Exception:
        pass
    finally:
        writer.close()


async def handle(client_r, client_w):
    try:
        head = await client_r.readuntil(b"\r\n\r\n")
    except Exception:
        client_w.close()
        return
    try:
        request_line = head.split(b"\r\n", 1)[0].decode()
        method, target, _ = request_line.split(" ", 2)
    except Exception:
        client_w.close()
        return
    try:
        up_r, up_w = await asyncio.open_connection(UP_HOST, UP_PORT)
    except Exception:
        client_w.close()
        return
    try:
        if method.upper() == "CONNECT":
            up_w.write(f"CONNECT {target} HTTP/1.1\r\nHost: {target}\r\n{UP_AUTH}\r\n".encode())
            await up_w.drain()
            resp = await up_r.readuntil(b"\r\n\r\n")
            if b" 200 " not in resp.split(b"\r\n", 1)[0]:
                client_w.write(b"HTTP/1.1 502 Bad Gateway\r\n\r\n")
                await client_w.drain()
                client_w.close(); up_w.close()
                return
            client_w.write(b"HTTP/1.1 200 Connection Established\r\n\r\n")
            await client_w.drain()
        else:
            # plain-HTTP: forward absolute URI through upstream
            lines = head.split(b"\r\n")
            lines[0] = f"{method} {target} HTTP/1.1".encode()
            out = b"\r\n".join(
                l for l in lines
                if not l.lower().startswith(b"proxy-authorization:"))
            out += UP_AUTH.encode() + b"\r\n"
            # strip any body already read? (GETs have none; keep simple)
            up_w.write(out)
            await up_w.drain()
        await asyncio.gather(relay(client_r, up_w), relay(up_r, client_w))
    finally:
        up_w.close()


async def main():
    srv = await asyncio.start_server(handle, *LISTEN)
    print(f"forward proxy on {LISTEN[0]}:{LISTEN[1]} -> {UP_HOST}:{UP_PORT}", flush=True)
    await srv.serve_forever()


if __name__ == "__main__":
    asyncio.run(main())
