from setuptools import setup
setup(name='ExcelToLaTeX',
    version='0.1',
	description='Python package to create LaTeX file from Excel data.',
	url='https://github.com/gknave/ExcelToLaTeX',
	author='Gary Nave',
	author_email='gknave@vt.edu',
	packages=['ExcelToLaTeX'],
	install_requires=[
	  'pandas',
	]
	zip_safe=False)