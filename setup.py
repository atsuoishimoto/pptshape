import os
from setuptools import setup

setup(
    name = "ppt-shape",
    version = "0.0.1",
    author = "Atsuo Ishimoto",
    author_email = "ishimoto@gembook.org",
    description = "Extract images from PPT files for Sphinx.",
    license = "BSD",
    url = "http://www.gembook.org",
    packages=['pptshape'],
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Topic :: Utilities",
        "License :: OSI Approved :: BSD License",
    ],
)
