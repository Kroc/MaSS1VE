Attribute VB_Name = "S1Z"
Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: S1Z

'MaSS1VE's project file format. This module reads S1Z files and populates the editor's _
 structures and can write the format out. "S1" stands for Sonic 1 and "Z" for ZIP; _
 the format is a simple ZIP file containing a bunch of files in various standardised _
 formats

