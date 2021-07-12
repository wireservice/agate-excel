#!/usr/bin/env python

from setuptools import setup

setup(
    name='agate-excel',
    version='0.2.4',
    description='agate-excel adds read support for Excel files (xls and xlsx) to agate.',
    long_description=open('README.rst').read(),
    author='Christopher Groskopf',
    author_email='chrisgroskopf@gmail.com',
    url='http://agate-excel.readthedocs.org/',
    license='MIT',
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Intended Audience :: Science/Research',
        'License :: OSI Approved :: MIT License',
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: Implementation :: CPython',
        'Programming Language :: Python :: Implementation :: PyPy',
        'Topic :: Scientific/Engineering :: Information Analysis',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
    packages=[
        'agateexcel'
    ],
    install_requires=[
        'agate>=1.5.0',
        'olefile',
        'openpyxl>=2.3.0',
        'xlrd>=0.9.4',
    ],
    extras_require={
        'test': [
            'nose>=1.1.2',
        ],
        'docs': [
            'Sphinx>=1.2.2',
            'sphinx_rtd_theme>=0.1.6',
        ],
    }
)
