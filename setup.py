import setuptools
#
with open("README.md", "r") as fh:
    long_description = fh.read();
#
setuptools.setup(
    name="pySW",
    packages = ['pySW'],
    version="1.0",
    author="Kalyan Inamdar",
    author_email="kalyaninamdar@protonmail.com",
    url = 'https://github.com/kalyanpi4/pySW',
    download_url='https://github.com/kalyanpi4/pySW/archive/V1.0.tar.gz',
    description="A Python wrapper around Solidworks built-in VBA API for modifying Solidworks assemblies and parts.",
    install_requires=[
          'numpy',
          'pandas',
          'pywin32',
    ],
    classifiers=[
		'License :: OSI Approved :: LGPL-2.0',
        'Programming Language :: Python :: 3',
    ],
    python_requires='>=3.6',
)