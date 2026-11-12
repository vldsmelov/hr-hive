"""Bounded local extraction workers. This module never opens a database."""
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
import os

from .common import sha
from .documents import extract_file


def _one(path, use_cache):
    os.environ['OMP_THREAD_LIMIT'] = '1'
    try:
        return str(path), extract_file(path, use_cache=use_cache), None
    except Exception as error:
        return str(path), None, f'{type(error).__name__}: {error}'


def iter_extracted(paths, max_workers=2, use_cache=True):
    """Yield (path, document_or_none, error_or_none); caller alone saves to DB.

    Identical bytes are extracted once. Cache writes use atomic unique files.
    Aliases retain their original filename while sharing the same page content.
    Prefer a regular filename over a download alias without reordering aliases.
    """
    if not 1 <= max_workers <= 2:
        raise ValueError('Допустимо не более двух процессов извлечения')
    grouped = {}
    for path in paths:
        path = str(Path(path))
        try:
            content_hash = sha(path)
        except OSError as error:
            yield path, None, f'{type(error).__name__}: {error}'
            continue
        grouped.setdefault(content_hash, []).append(path)
    if not grouped:
        return
    with ProcessPoolExecutor(max_workers=max_workers) as pool:
        tasks = {}
        for names in grouped.values():
            preferred = next((name for name in names
                              if Path(name).suffix.lower() != '.crdownload'), names[0])
            tasks[pool.submit(_one, preferred, use_cache)] = names
        for task in as_completed(tasks):
            names = tasks[task]
            _, document, error = task.result()
            for name in names:
                result = {**document, 'original_name': Path(name).name} if document else None
                yield name, result, error
