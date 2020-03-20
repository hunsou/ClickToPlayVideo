@echo off
copy %~dp0\mscomm32.ocx c:\
regsvr32 c:\mscomm32.ocx