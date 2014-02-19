@ECHO OFF

REM rename.bat: execute rename.vbs
REM Copyright (C) 2014 ITIIC <http://itiic.com/>
REM
REM rename.vbs is free software: you can redistribute it and/or modify
REM it under the terms of the GNU Lesser General Public License as published by
REM the Free Software Foundation, either version 3 of the License, or
REM (at your option) any later version.
REM
REM INIFile.vbs is distributed in the hope that it will be useful,
REM but WITHOUT ANY WARRANTY; without even the implied warranty of
REM MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
REM GNU Lesser General Public License for more details.
REM
REM You should have received a copy of the GNU Lesser General Public License
REM along with INIFile.vbs. If not, see <http://www.gnu.org/licenses/>.
REM
REM --------------------------------------------------------------------
REM
REM Usage: rename.bat <ENTER>
REM
REM Put rename.bat, rename.vbs, exiv2.exe in dir with jpg's files

cscript /nologo rename.vbs
