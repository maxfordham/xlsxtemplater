#!/usr/bin/env python

"""The setup script."""

from setuptools import setup, find_packages
import versioneer

def requirements():
    with open('requirements.txt') as requirements_file:
        _requirements = requirements_file.readlines()
    return _requirements

def readme():
    with open('README.md') as f:
        README = f.read()
    return README

setup(
    author="John Gunstone",
    author_email='j.gunstone@maxfordham..com',
    python_requires='>=3.5',
    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Natural Language :: English',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
    ],
    description="high-level wrapper that sits on top of xlsxwriter,\
        to template the output of the pandas dataframes to formatted excel tables",
    install_requires=requirements(),
    license="BSD license",
    long_description=readme(),
    long_description_content_type="text/markdown",
    include_package_data=True,
    keywords='xlsxtemplater',
    name='xlsxtemplater',
    packages=find_packages(include=['xlsxtemplater', 'xlsxtemplater.*']),
    url='https://github.com/gunstonej/xlsxtemplater',
    version=versioneer.get_version(),
    cmdclass=versioneer.get_cmdclass(),
)
