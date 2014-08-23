import os
from setuptools import setup

def read(fname):
    return open(
            os.path.join(os.path.dirname(__file__), fname)).read()

setup(
    name = "pptshape",
    version = "0.0.4",
    author = "Atsuo Ishimoto",
    author_email = "ishimoto@gembook.org",
    description = "Extract images from PowerPoint presentation files for Sphinx.",
    license = "BSD",
    url = "https://github.com/atsuoishimoto/pptshape",
    long_description=read('README.rst'),
    install_requires=[
        'sphinx',
        ],
    packages=['pptshape'],
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Topic :: Software Development :: Documentation",
        "License :: OSI Approved :: BSD License",
        "Programming Language :: Python :: 2.7",
        "Programming Language :: Python :: 3",
    ],
)
