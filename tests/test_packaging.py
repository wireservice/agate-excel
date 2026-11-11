import shutil
import tarfile
import zipfile
from pathlib import Path

from build import ProjectBuilder


def test_wheel_contains_only_library_packages(tmp_path):
    source = tmp_path / 'source'
    shutil.copytree(
        Path(__file__).resolve().parents[1], source,
        ignore=shutil.ignore_patterns('.git', 'build', 'dist', '*.egg-info', '__pycache__', '.pytest_cache'),
    )
    distributions = tmp_path / 'dist'
    sdist = ProjectBuilder(source).build('sdist', distributions)
    with tarfile.open(sdist) as archive:
        names = archive.getnames()
        for directory in ('docs', 'examples', 'tests'):
            assert any(f'/{directory}/' in name for name in names)
        archive.extractall(tmp_path / 'unpacked')
    unpacked_source = next((tmp_path / 'unpacked').iterdir())
    for root in (source, unpacked_source):
        wheel = ProjectBuilder(root).build('wheel', distributions)
        with zipfile.ZipFile(wheel) as archive:
            names = archive.namelist()
            assert 'agateexcel/__init__.py' in names
            assert all(name.startswith('agateexcel/') or '.dist-info/' in name for name in names)
