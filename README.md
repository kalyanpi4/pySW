# pySW
[![DOI](https://zenodo.org/badge/237943198.svg)](https://zenodo.org/badge/latestdoi/237943198)
A Wrapper around Solidworks VBA API for Automating Geometry Modifications for Python-based Optimization and Design Space Exploration

pySW is simply a Python wrapper around Solidworks built-in VBA API for automated modifications to Solidworks assemblies and parts.

Primary purpose of pySW is optimization studies, Design Space Exploration studies etc. There are many good libraries which offer framework for single and multi-ojective optimization, like [pymoo](https://pymoo.org/index.html), [openMDAO](https://openmdao.org/) and [pyOpt](http://www.pyopt.org/) written in Python itself. Design space can be explored using libraries like [pyDOE](https://pythonhosted.org/pyDOE/index.html). Apart from these, the famous [Scipy](https://docs.scipy.org/doc/scipy/reference/optimize.html) library provides functions for minimizing (or maximizing) objective functions, possibly subject to constraints.

In many cases the optimization or the space exploration problem is not straight-forward that it can be expressed as equations. Some problems require some communication link between various software. For example, consider an optimization of a winglet of a commercial aircraft. This problem requires modifying and saving geometry using a CAD program, using a CAE program for analysing and saving results of the analysis and a third program/code to act as a link between CAD and CAE code as well as perform the tasks of optimization or space exploration. The first task can be prformed using pySW and the third task using the libraries mentioned above.

Note: Solidworks is a proprietary software of 3DS Systems. If you hav access to Solidworks, pySW will make your life much easier if you want to optimize using Solidworks.
As an option, FreeCAD is an open-source primarily CAD program written completely in Python. It also as modules for CFD using OpenFOAM and FEM analysis.

It currently does not have the capability to modify or create sketches. 


### Installation

You can install pySW from pip from the command prompt by running:

```sh
pip install pySW
```

Dependencies:
1. [pywin32](https://pypi.org/project/pywin32/)
2. [Numpy](https://numpy.org/)
3. [Pandas](https://pandas.pydata.org/)

You can install the dependencies by running the following commands in the command prompt. If you have installed the open-source Anaconda distribution for Python, please check if the libaries are already installed.
```sh
pip install pywin32
pip install numpy
pip install pandas
```

License
----

GNU Lesser General Public License v2


[//]: # (These are reference links used in the body of this note and get stripped out when the markdown processor does its job. There is no need to format nicely because it shouldn't be seen. Thanks SO - http://stackoverflow.com/questions/4823468/store-comments-in-markdown-syntax)


   [dill]: <https://github.com/joemccann/dillinger>
   [pymoo]: <https://pymoo.org/index.html>
   [openMDAO]: <https://openmdao.org/>
   [pyOpt]: <http://www.pyopt.org/>
   [pyDOE]: <https://pythonhosted.org/pyDOE/index.html>
   [Scipy]: <https://docs.scipy.org/doc/scipy/reference/optimize.html>
   [pywin32]: <https://pypi.org/project/pywin32/>
   [Numpy]: <https://numpy.org/>
   [pandas]: <https://pandas.pydata.org/>
