#!/usr/bin/env python3
"""Bridge a stdio MCP client to Office Studio's local, read-only HTTP MCP server.

Standard library only. Run the app first, then configure an MCP client to launch:
  python /absolute/path/examples/plugins/mcp_stdio.py
Optional: --url http://127.0.0.1:3001/mcp
"""
import argparse
import json
import sys
import urllib.error
import urllib.parse
import urllib.request


def main():
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--url', default='http://127.0.0.1:3001/mcp')
    args = parser.parse_args()
    url = urllib.parse.urlparse(args.url)
    if url.scheme != 'http' or url.hostname not in ('127.0.0.1', 'localhost', '::1') or url.path != '/mcp' or url.username or url.password:
        parser.error('--url must be the local HTTP /mcp endpoint')
    protocol_version = '2025-11-25'
    # Do not route local workspace content through environment-configured proxies.
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    for line in sys.stdin:
        message = None
        response = None
        try:
            message = json.loads(line)
            request = urllib.request.Request(args.url, data=json.dumps(message, ensure_ascii=False).encode('utf-8'), headers={
                'Content-Type': 'application/json', 'Accept': 'application/json, text/event-stream', 'MCP-Protocol-Version': protocol_version,
            }, method='POST')
            with opener.open(request, timeout=15) as result:
                payload = result.read()
            if payload:
                response = json.loads(payload)
                if isinstance(message, dict) and message.get('method') == 'initialize' and 'result' in response:
                    protocol_version = response['result']['protocolVersion']
        except urllib.error.HTTPError as error:
            try:
                response = json.loads(error.read())
            except (ValueError, OSError):
                response = {'jsonrpc': '2.0', 'id': message.get('id') if isinstance(message, dict) else None, 'error': {'code': -32603, 'message': f'Local server returned HTTP {error.code}.'}}
        except json.JSONDecodeError:
            response = {'jsonrpc': '2.0', 'id': None, 'error': {'code': -32700, 'message': 'Invalid JSON.'}}
        except (OSError, ValueError) as error:
            response = {'jsonrpc': '2.0', 'id': message.get('id') if isinstance(message, dict) else None, 'error': {'code': -32603, 'message': 'Cannot connect to Office Studio. Start npm run dev first.'}}
            print(str(error), file=sys.stderr, flush=True)
        # Never reply to valid notifications, including notifications/initialized.
        if response is not None and (not isinstance(message, dict) or 'id' in message or message.get('jsonrpc') != '2.0'):
            print(json.dumps(response, ensure_ascii=False, separators=(',', ':')), flush=True)


if __name__ == '__main__':
    main()
