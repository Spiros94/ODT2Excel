from distutils.core import setup
import py2exe, sys, os

setup(
	console = [{
		"script":"odt2excel.py"
        }],
        zipfile = None,
        data_files = None,
        version = "1.0",
        name = "Odt To Xlsx",
        description = "A converter program for odt files",
        author = "Mitropoulos Spiros",
        license = "GNU GPL 3 License",
        url = "http://www.eyrhka.gr",
)
