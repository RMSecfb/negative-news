"""Read/write one atomic report revision; optimistic retries preserve concurrent manual edits."""
from __future__ import annotations
import base64
import io
import json
import time
from pathlib import Path
import requests


class RepositoryError(RuntimeError):
    pass


class GitHubRepository:
    def __init__(self, repo, token='', branch='main'):
        self.repo, self.token, self.branch = repo, token, branch
        self.base = f'https://api.github.com/repos/{repo}'
        self.session = requests.Session()
        self.session.headers.update({'Accept': 'application/vnd.github+json'})
        if token:
            self.session.headers['Authorization'] = 'Bearer ' + token

    def api(self, method, path, **kwargs):
        response = self.session.request(method, self.base + path, timeout=60, **kwargs)
        if response.status_code >= 400:
            raise RepositoryError(f'GitHub HTTP {response.status_code}：{method} {path}')
        return response.json() if response.content else {}

    def revision(self):
        return self.api('GET', '/git/ref/heads/' + self.branch)['object']['sha']

    def read(self, path, revision=None):
        response = self.session.get(self.base + '/contents/' + path,
                                    params={'ref': revision or self.branch}, timeout=45)
        if response.status_code == 404:
            return None
        if response.status_code != 200:
            raise RepositoryError(f'讀取 {path} 失敗：GitHub HTTP {response.status_code}')
        payload = response.json()
        if payload.get('encoding') == 'base64':
            return base64.b64decode(payload.get('content', ''))
        blob = self.api('GET', '/git/blobs/' + payload['sha'])
        return base64.b64decode(blob['content'])

    def list(self, prefix, revision=None):
        tree = self.api('GET', '/git/trees/' + (revision or self.revision()) + '?recursive=1')
        if tree.get('truncated'):
            raise RepositoryError('GitHub 目錄清單不完整，已停止以免漏讀資料')
        return [r['path'] for r in tree.get('tree', []) if r['type'] == 'blob' and r['path'].startswith(prefix)]

    def transaction(self, build, message):
        if not self.token:
            raise RepositoryError('需要既有 GITHUB_TOKEN 的 Contents: Read and write 權限才能儲存')
        for attempt in range(3):
            revision = self.revision()
            files, result = build(lambda path: self.read(path, revision), revision)
            if not files:
                return result
            tree = self.api('GET', '/git/commits/' + revision)['tree']['sha']
            entries = []
            for path, data in files.items():
                raw = data.encode('utf-8') if isinstance(data, str) else data
                sha = self.api('POST', '/git/blobs', json={'content': base64.b64encode(raw).decode(), 'encoding': 'base64'})['sha']
                entries.append({'path': path, 'mode': '100644', 'type': 'blob', 'sha': sha})
            new_tree = self.api('POST', '/git/trees', json={'base_tree': tree, 'tree': entries})['sha']
            commit = self.api('POST', '/git/commits', json={'message': message, 'tree': new_tree, 'parents': [revision]})['sha']
            response = self.session.patch(self.base + '/git/refs/heads/' + self.branch,
                                          json={'sha': commit, 'force': False}, timeout=45)
            if response.status_code == 200:
                return result
            if response.status_code not in (409, 422):
                raise RepositoryError(f'提交失敗：GitHub HTTP {response.status_code}')
            time.sleep(1 + attempt)
        raise RepositoryError('資料同時被其他執行更新；已停止提交，請重新載入後再試')


class LocalRepository:
    """Same transaction contract for tests and offline verification."""
    def __init__(self, root):
        self.root = Path(root)

    def read(self, path, revision=None):
        target = self.root / path
        return target.read_bytes() if target.exists() else None

    def list(self, prefix, revision=None):
        return [p.relative_to(self.root).as_posix() for p in self.root.rglob('*')
                if p.is_file() and p.relative_to(self.root).as_posix().startswith(prefix)]

    def transaction(self, build, message):
        files, result = build(self.read, 'local')
        for path, raw in files.items():
            target = self.root / path
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(raw.encode('utf-8') if isinstance(raw, str) else raw)
        return result


def read_json(raw, default=None):
    return json.loads(raw.decode('utf-8-sig')) if raw else (default if default is not None else {})


def json_bytes(value):
    return json.dumps(value, ensure_ascii=False, indent=2, default=str).encode('utf-8')
