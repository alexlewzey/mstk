#!/usr/bin/env python

"""The setup script."""

from setuptools import setup, find_packages

with open('README.rst') as readme_file:
    readme = readme_file.read()

setup(
    author="Alexander Lewzey",
    author_email='a.lewzey@hotmail.co.uk',
    python_requires='>=3.5',
    description="A collection of general purpose helper modules",
    entry_points={
        'console_scripts': [
            'mstk=mstk.cli:main',
        ],
    },
    install_requires=[
        'pandas',
        'python-docx',
        'python-pptx',
    ],
    license="BSD license",
    keywords='mstk',
    name='mstk',
    packages=find_packages(include=['mstk', 'mstk.*']),
    test_suite='tests',
    url='https://github.com/alexlewzey/mstk',
    version='0.1.0',
)
