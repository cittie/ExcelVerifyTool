# -*-code: utf-8-*-
from setuptools import setup, find_packages

install_requires = [
    "xlrd",
    "BeautifulSoup4"
    ]
setup(
    name = 'VerifyTool',
    version = '0.0.1',
    packages = find_packages(exclude=['*.pyc']),
    author="Studio.QA.Beijing",
    author_email="Studio.QA.Beijing@glu.com",
    platforms="ANY",
    install_requires=install_requires,
)
