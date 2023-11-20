from setuptools import find_packages, setup

with open('README.rst') as f:
    long_description = f.read()

setup(
    name='agate-excel',
    version='0.4.1',
    description='agate-excel adds read support for Excel files (xls and xlsx) to agate.',
    long_description=long_description,
    long_description_content_type='text/x-rst',
    author='Christopher Groskopf',
    author_email='chrisgroskopf@gmail.com',
    url='https://agate-excel.readthedocs.org/',
    license='MIT',
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Intended Audience :: Science/Research',
        'License :: OSI Approved :: MIT License',
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
        'Programming Language :: Python :: Implementation :: CPython',
        'Programming Language :: Python :: Implementation :: PyPy',
        'Topic :: Scientific/Engineering :: Information Analysis',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
    packages=find_packages(exclude=['tests', 'tests.*']),
    install_requires=[
        'agate>=1.5.0',
        'olefile',
        'openpyxl>=2.3.0',
        'xlrd>=0.9.4',
    ],
    extras_require={
        'test': [
            'pytest',
            'pytest-cov',
        ],
    }
)
