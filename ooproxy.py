#!/usr/bin/python3
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

from io import BytesIO
import uno
import unohelper
import json
import eventlet
import argparse
import logging
import logging.handlers
from contexttimer import Timer
from contexttimer import timer
from eventlet import Timeout
from com.sun.star.beans import PropertyValue
from com.sun.star.uno import Exception as UnoException
from com.sun.star.connection import NoConnectException, ConnectionSetupException
from com.sun.star.beans import UnknownPropertyException
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.io import XOutputStream
from com.sun.star.io import IOException

_logger = logging.getLogger("ooproxy")
 
class OutputStreamWrapper(unohelper.Base, XOutputStream):
    """ Minimal Implementation of XOutputStream """
    def __init__(self, outstream=None):
        self.data = outstream or BytesIO()
 
    def writeBytes(self, b):
        self.data.write(b.value)
 
    def close(self):
        self.data.close()
         
    def flush(self):
        self.data.flush()
        
    def closeOutput(self):
        pass
    
    
class TimeoutException(Exception):
    pass


class NoDataExeption(Exception):
    pass


def info(message):
    _logger.info(message)

def toProperties(**args):
    props = []
    for key in args:
        prop = PropertyValue()
        prop.Name = key
        prop.Value = args[key]
        props.append(prop)
    return tuple(props)    

def application(fd, sock, args):
    # vars
    peer_name = repr(new_sock.getpeername())
    timeout = args.timeout
    header = None
    document = None
    
    # helpers functions
    
    def writeln(resp):
        line = "%s\n" % resp
        line = line.encode(encoding="utf-8")
        fd.write(line)
    
    def readHeader():
        with Timeout(timeout,TimeoutException):
            line = fd.readline()
            if not line:
                raise NoDataExeption()
            line = line.decode("utf-8")
            return json.loads(line)
        
    def readData():
        with Timeout(timeout,TimeoutException):
            res = fd.read(header["length"])
            if not res:
                raise NoDataExeption()
            return res
            
    def closeDocument():
        try:
            if document:
                document.close()
        except:
            pass
        finally:
            document=None
   
    def refreshDocument():
        # At first update Table-of-Contents.
        # ToC grows, so page numbers grows too.
        # On second turn update page numbers in ToC
        try:
            document.refresh()
            indexes = document.getDocumentIndexes()
            for i in range(0, indexes.getCount()):
                indexes.getByIndex(i).update()
        except AttributeError: # ods document does not support refresh
            # the document doesn't implement the XRefreshable and/or
            # XDocumentIndexesSupplier interfaces
            pass

    # logic
    
    try:
        header = readHeader()
        
        # initialize     
        host = header.get("host",args.oo_host)
        port = header.get("port",args.oo_port)
        timeout = header.get("timeout",timeout)
         
        ooLocalCtx = uno.getComponentContext()        
        ooLocalResolver = ooLocalCtx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ooLocalCtx)
    
        ooRemoteCtx = ooLocalResolver.resolve("uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % (host, port))
        ooRemoteServiceManager = ooRemoteCtx.ServiceManager
        ooRemoteDesktop = ooRemoteServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ooRemoteCtx)
        
        # ok        
        writeln("{}")
        while True:
            # read
            header = readHeader()
            fnct = header["fnct"]
            if fnct == "close":
                break
            elif fnct == "closeDocument":
                closeDocument()
                writeln('{}')
            elif fnct == "putDocument":
                # cleanup
                closeDocument()
                # read data
                with Timer() as t:
                    data = readData()
                    inputStream = ooRemoteServiceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", ooRemoteCtx)
                    inputStream.initialize((uno.ByteSequence(data),))
                    _logger.debug("read data takes %s" % t.elapsed)
                    
                # load document
                with Timer() as t:
                    try:
                        document = ooRemoteDesktop.loadComponentFromURL('private:stream', "_blank", 0, toProperties(InputStream = inputStream))
                    finally:
                        inputStream.closeInput()
                    _logger.debug("load document takes %s" % t.elapsed)
                writeln('{}')
            elif fnct == "printDocument":
                printer = header.get("printer")
                if printer:
                    uno.invoke(document,"setPrinter", toProperties(Name = printer) )
                uno.invoke(document, "print", toProperties(Wait = True) )
                writeln('{}')      
            elif fnct == "refreshDocument":
                refreshDocument()
                writeln('{}')
            elif fnct == "getDocument":
                filter_name = header.get("filter")
                refreshDocument()
                out = OutputStreamWrapper()
                try:
                    document.storeToURL("private:stream", toProperties(OutputStream = out, FilterName = filter_name))
                    writeln('{"length" : %s }' % len(out.data.getvalue()))
                    fd.write(out.data.getvalue())                   
                except IOException:
                    writeln('{ "error" : "io-error", "message" : "Exception during conversion" }')
            elif fnct == "streamDocument":
                with Timer() as t:
                    filter_name = header.get("filter")
                    refreshDocument()
                    out = OutputStreamWrapper(fd)   
                    document.storeToURL("private:stream", toProperties(OutputStream = out, FilterName = filter_name))
                    _logger.debug("streamDocument takes %s" % t.elapsed)
                break # break after stream                
            elif fnct == "insertDocument":
                data = readData()
                placeholder_text = "<insert_doc('%s')>" % header["name"]
                
                # prepare stream
                inputStream = ooRemoteServiceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", ooRemoteCtx)
                inputStream.initialize((uno.ByteSequence(data),))
                
                # search
                search = document.createSearchDescriptor()
                search.SearchString = placeholder_text
                found = document.findFirst(search)
                
                # insert
                error = False
                while found:
                    try:
                        found.insertDocumentFromURL('private:stream', toProperties(InputStream = inputStream, FilterName = "writer8"))
                        error = False
                    except Exception:
                        error = True
                if error:
                    writeln('{ "error" : "io-error", "message" : "Unable to insert document %s" }' % header["name"])
                else:
                    writeln("{}")
                    
            elif fnct == "addDocument":
                data = readData()
                inputStream = ooRemoteServiceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", ooRemoteCtx)
                inputStream.initialize((uno.ByteSequence(data),))
                try:
                    document.Text.getEnd().insertDocumentFromURL('private:stream', toProperties(InputStream = inputStream, FilterName = "writer8"))
                    writeln("{}")
                except Exception:
                    writeln('{ "error" : "io-error", "message" : "Unable to insert document" }')
                    
            fd.flush()
                 
    except IllegalArgumentException:
        writeln('{ "error": "invalid-url", "message" : "The url is invalid (%s)"')
    except NoConnectException:
        writeln('{ "error": "no-connection", "message" : "Failed to connect to OpenOffice.org on host %s, port %s"' % (host, port))
    except ConnectionSetupException:
        writeln('{ "error": "no-access", "message" : "Not possible to accept on a local resource"')
    except TimeoutException:
        info("Socket Timeout %s" % peer_name)
    except NoDataExeption:
        info("Stream closed %s" % peer_name)
    except Exception as e:
        _logger.fatal(e)
        writeln('{ "error": "unexpected", "message" : "Unexpected error"')        
    finally:        
        closeDocument()    
        fd.close()
        sock.close()
        info("Client %s disconnected" % peer_name)
    

if __name__ == '__main__':
    logging.basicConfig()
    logging.getLogger().setLevel(logging.DEBUG)
    
    parser = argparse.ArgumentParser(description="OOProxy Arguments")
    parser.add_argument("--oo-host", metavar="OO_HOST", type=str, help="The OpenOffice Server Host", default="127.0.0.1")
    parser.add_argument("--oo-port", metavar="OO_PORT", type=int, help="The OpenOffice Server Port", default=8100)
    parser.add_argument("--port", metavar="PORT", type=int, help="The OOProxy Port", default=8099)
    parser.add_argument("--listen", metavar="LISTEN", type=str, help="Listen on Address", default="127.0.0.1")
    parser.add_argument("--timeout", metavar="TIMEOUT", type=int, help="Timeout in Seconds", default=5)
    parser.add_argument("--bufsize", metavar="BUFSIZE", type=int, help="Buffer size", default=32768)
    parser.add_argument("--syslog", dest='syslog', action="store_true", help="Log to syslog")
    args = parser.parse_args()
    
    if args.syslog:
        syslog_handler = logging.handlers.SysLogHandler(address="/dev/log")
        syslog_handler.setFormatter(logging.Formatter('%(name)s [%(levelname)s] %(message)s'))
        logging.getLogger().addHandler(syslog_handler)
    
    info("Server listen on port %s" % args.port)
    server = eventlet.listen((args.listen, args.port))
    pool = eventlet.GreenPool()
    while True:
        try:
            new_sock, address = server.accept()
            _logger.info("Client %s connected" % repr(new_sock.getpeername()))
            fd = new_sock.makefile("rw")
            pool.spawn_n(application, fd, new_sock, args)
        except (SystemExit, KeyboardInterrupt):
            break
    #from werkzeug.serving import run_simple
    #run_simple('localhost', DEFAUL_DEFAULT_PROXY_PORT, application)

