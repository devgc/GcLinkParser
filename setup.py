# Tool to parse Link Files and Jump Lists
#
# Copyright (C) 2015, G-C Partners, LLC <dev@g-cpartners.com>
#
# Refer to AUTHORS for acknowledgements.
#
# This software is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This software is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with this software.  If not, see <http://www.gnu.org/licenses/>.

from distutils.core import setup
import py2exe, sys, os

sys.argv.append('py2exe')

setup(
	options = {'py2exe': {'bundle_files': 1}},
	console  = [{'script': "GcLinkParser.py"}],
	zipfile = None,
)