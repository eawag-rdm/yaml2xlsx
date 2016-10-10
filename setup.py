# _*_ coding: utf-8 _*_

from setuptools import setup

setup(
    name='yaml2xlsx',
    version='1.0.0',
    description='Converts a file containing YAML-documents to an xlsx worksheet',
    url='https://github.com/eawag-rdm/yaml2xlsx.git',
    author='Harald von Waldow',
    author_email='harald.vonwaldow@eawag.ch',
    license='GNU AFFERO GENERAL PUBLIC LICENSE Version 3',
    classifiers = [
        'Development Status :: 4 - Beta'],
    keywords = ['OOXML', 'xlsx', 'YAML'],
    install_requires = ['PyYaml', 'openpyxl'],
    packages = ['test', 'yaml2xlsx'],
    entry_points = {
        'console_scripts': [
            'yaml2xlsx=yaml2xlsx.yaml2xlsx:main']}
    )
