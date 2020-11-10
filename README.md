# Codfusion: pythonforword
Is nothing more than a service to use within applications wich generates a Word-document with the use of a template-file.

**NOTE: all done on a Windows machine.**

## What you need to do to get it 'up and running'?
First of all you have to install Python on your machine. In this case I installed Python 3.8.5.

After installing Python, install two depencies: 'docx' and 'docxtpl'. In order to install both dependencies, use the command prompt to navigate to the pip directory whitin the Python directory and run the following commands

``` bash
pip install docx

pip install docxtpl
```
That's it.

## How to use
The service has 2 public functions: 'init' and 'genWordDoc'.

The 'init' function is to setup all required settings before running 'genWordDoc'.

**NOTE: change the argument 'pythonExePath' with the actual path to Python** 

It's important to run the 'init' first.
