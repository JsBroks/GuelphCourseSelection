
import os
import pip


if not os.path.exists("out"):
    os.makedirs("out")

if not os.path.exists("html_courses"):
    os.makedirs("html_courses")

pkgs = ['xlsxwriter', 'beautifulsoup4']
for package in pkgs:
    pip.main(['install', package])
