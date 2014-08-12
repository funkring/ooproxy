# -*- coding: utf-8 -*-
#############################################################################
#
#    Copyright (c) 2007 Martin Reisenhofer <martin.reisenhofer@funkring.net>
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################

from werkzeug.wrappers import Request, Response

import uno
import unohelper
from com.sun.star.beans import PropertyValue
from com.sun.star.uno import Exception as UnoException
from com.sun.star.connection import NoConnectException, ConnectionSetupException
from com.sun.star.beans import UnknownPropertyException
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.io import XOutputStream
from com.sun.star.io import IOException
 
DEFAUL_DEFAULT_OPENOFFICE_PORT = 8100
DEFAUL_DEFAULT_PROXY_PORT = 8099

@Request.application
def application(request):
    return Response('Hello World!')

if __name__ == '__main__':
    from werkzeug.serving import run_simple
    run_simple('localhost', DEFAUL_DEFAULT_PROXY_PORT, application)

