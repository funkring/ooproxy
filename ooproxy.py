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
import logging.handlers
from contexttimer import Timer
from eventlet import Timeout
from com.sun.star.beans import PropertyValue
from com.sun.star.uno import Exception as UnoException
from com.sun.star.connection import NoConnectException, ConnectionSetupException
from com.sun.star.beans import UnknownPropertyException
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.io import XOutputStream
from com.sun.star.io import IOException
from com.sun.star.io import NotConnectedException

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

class UnsupportedException(Exception):
    def __init__(self, message):
        self.message = message

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

class OOProxy(object):
    def __init__(self, fd, sock, args):
        self.fd = fd
        self.sock = sock
        self.args = args
        self.peer_name = repr(sock.getpeername())
        self.header = None
        self.document = None
        self.ooLocalResolver = None
        self.ooRemoteDesktop = None
        self.timeout = args.timeout
        self.host = args.oo_host
        self.port = args.oo_port
        
    # helpers functions
    def writeln(self,resp):
        line = "%s\n" % resp
        line = line.encode(encoding="utf-8")
        self.fd.write(line)
        self.fd.flush();
    
    def readHeader(self):
        with Timeout(self.timeout,TimeoutException):
            line = self.fd.readline()
            if not line:
                raise NoDataExeption()
            line = line.decode("utf-8")
            self.header = json.loads(line)
        
    def readData(self):
        with Timeout(self.timeout,TimeoutException):
            _logger.info("Read %s bytes of Data" % self.header["length"])
            res = self.fd.read(self.header["length"])
            if not res:
                raise NoDataExeption()
            return res
        
    def closeDocument(self):
        try:
            if self.document:
                self.document.close(False)
        except Exception as e:
            _logger.error(e)
            
        try:
            if self.document:
                self.document.dispose()
        except Exception as e:
            _logger.error(e)
            
        self.document=None
            
    def cleanup(self):
        self.closeDocument()
       
    def refreshDocument(self):
        # At first update Table-of-Contents.
        # ToC grows, so page numbers grows too.
        # On second turn update page numbers in ToC
        try:
            self.document.refresh()
            indexes = self.document.getDocumentIndexes()
            for i in range(0, indexes.getCount()):
                indexes.getByIndex(i).update()
        except AttributeError: # ods document does not support refresh
            # the document doesn't implement the XRefreshable and/or
            # XDocumentIndexesSupplier interfaces
            pass
        
    def run(self):
        try:
            # read first
            self.readHeader()
        
            # initialize     
            self.host = self.header.get("host",self.args.oo_host)
            self.port = self.header.get("port",self.args.oo_port)
            self.timeout = self.header.get("timeout",self.args.timeout)
    
            # open openoffice        
            self.ooLocalCtx = uno.getComponentContext()        
            self.ooLocalResolver = self.ooLocalCtx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", self.ooLocalCtx)
            self.ooRemoteCtx = self.ooLocalResolver.resolve("uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % (self.host, self.port))
            self.ooRemoteServiceManager = self.ooRemoteCtx.ServiceManager
            self.ooRemoteDesktop = self.ooRemoteServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ooRemoteCtx)
      
            # send ok        
            self.writeln("{}")
            
            # loop while run=true
            run = True
            while run:
                # read
                self.readHeader()
                fnct = self.header["fnct"]
                if fnct == "close":
                    run = False
                    self.writeln('{}')                    
                elif fnct == "closeDocument":
                    self.closeDocument()
                    self.writeln('{}')
                elif fnct == "putDocument":
                    # cleanup
                    self.closeDocument()

                    # read data
                    with Timer() as t:
                        data = self.readData()
                        _logger.debug("read data takes %s" % t.elapsed)
                        
                    # load document
                    with Timer() as t:
                        inputStream = self.ooRemoteServiceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", self.ooRemoteCtx)
                        try:
                            inputStream.initialize((uno.ByteSequence(data),))
                            self.document = self.ooRemoteDesktop.loadComponentFromURL('private:stream', "_blank", 0, toProperties(InputStream = inputStream))
                        finally:
                            inputStream.closeInput()
                        _logger.debug("load document takes %s" % t.elapsed)
                    self.writeln('{}')
                    
                elif fnct == "printDocument":
                    printer = self.header.get("printer")
                    if printer:
                        uno.invoke(self.document,"setPrinter", toProperties(Name = printer) )
                    uno.invoke(self.document, "print", toProperties(Wait = True) )
                    self.writeln('{}')
                          
                elif fnct == "refreshDocument":
                    self.refreshDocument()
                    self.writeln('{}')
                    
                elif fnct == "getDocument":
                    filter_name = self.header.get("filter")
                    self.refreshDocument()
                    out = OutputStreamWrapper()
                    try:
                        self.document.storeToURL("private:stream", toProperties(OutputStream = out, FilterName = filter_name))
                        data_len = len(out.data.getvalue())
                        self.writeln('{"length" : %s }' % data_len )
                        _logger.info("Send document with length=%s" % data_len)
                        self.fd.write(out.data.getvalue())                                               
                    except IOException:
                        self.writeln('{ "error" : "io-error", "message" : "Exception during conversion" }')
                    finally:
                        out.closeOutput()
                        
                elif fnct == "streamDocument":
                    run = False
                    with Timer() as t:
                        filter_name = self.header.get("filter")
                        self.refreshDocument()
                        out = OutputStreamWrapper(self.fd)   
                        try:
                            self.document.storeToURL("private:stream", toProperties(OutputStream = out, FilterName = filter_name))
                            _logger.debug("streamDocument takes %s" % t.elapsed)
                        finally:      
                            out.closeOutput()
                                
                elif fnct == "insertDocument":
                    data = self.readData()
                    name = self.header["name"]                    
                    placeholder_text = "<insert_doc('%s')>" % name
                    
                    filterName = "writer8"
                    if name.lower().endswith(".html"):
                        filterName = "HTML"
                    
                    # prepare stream
                    inputStream = self.ooRemoteServiceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", self.ooRemoteCtx)
                    try:
                        inputStream.initialize((uno.ByteSequence(data),))
                        
                        # search
                        search = self.document.createSearchDescriptor()
                        search.SearchString = placeholder_text
                        found = self.document.findFirst(search)
                        
                        # insert
                        error = None
                        while found:
                            try:
                                found.insertDocumentFromURL('private:stream', toProperties(InputStream = inputStream, FilterName = filterName))
                                found = self.document.findNext(found.End, search)
                            except Exception as e:
                                error = e.message or "Error"
                                break
                                
                        if error:
                            self.writeln('{ "error" : "io-error", "message" : "Unable to insert document %s, because %s" }' % (self.header["name"]),error)
                        else:
                            self.writeln("{}")
                    finally:
                        inputStream.closeInput()
                        
                elif fnct == "addDocument":
                    data = self.readData()
                    inputStream = self.ooRemoteServiceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", self.ooRemoteCtx)                    
                    try:
                        inputStream.initialize((uno.ByteSequence(data),))
                        self.document.Text.getEnd().insertDocumentFromURL('private:stream', toProperties(InputStream = inputStream, FilterName = "writer8"))
                        self.writeln("{}")
                    except Exception:
                        inputStream.closeInput()
                        self.writeln('{ "error" : "io-error", "message" : "Unable to insert document" }')
                else:
                    raise UnsupportedException(fnct)
                
                # FLUSH after finished loop                   
                self.fd.flush()
                
        except IllegalArgumentException:
            _logger.error("Invalid URL!")
            self.writeln('{ "error": "invalid-url", "message" : "The url is invalid }"')            
        except NoConnectException:
            _logger.error("Not connected!")
            self.writeln('{ "error": "no-connection", "message" : "Failed to connect to OpenOffice.org on host %s, port %s" }' % (self.host, self.port))
        except NotConnectedException:
            _logger.error("Failed accessing stream!")
            self.writeln('{ "error": "not-connected", "message" : "Connection to OpenOffice.org on host %s, port %s are lost" }' % (self.host, self.port))
        except ConnectionSetupException:
            self.writeln('{ "error": "no-access", "message" : "Not possible to accept on a local resource" }')
        except TimeoutException:
            info("Socket Timeout %s" % self.peer_name)
        except NoDataExeption:
            info("Stream closed %s" % self.peer_name)
        except UnsupportedException as e:
            msg = "Unsupported function %s" % e.message
            _logger.fatal(msg)
            self.writeln('{ "error": "unsupported", "message" : "%s" }' % msg)
        except Exception as e:
            _logger.exception(e)
            self.writeln('{ "error": "unexpected", "message" : "Unexpected error" }')        
        finally:
            try:
                self.fd.close()
            except Exception as e:
                _logger.exception(e)
                
            try:
                self.sock.close()
            except Exception as e:
                _logger.exception(e)
                
            try:
                self.cleanup()
            except:    
                _logger.exception(e)

            info("Client %s disconnected" % self.peer_name)
       
    
def application(client_fd, client_sock, args):
    OOProxy(client_fd, client_sock, args).run()
    
    
if __name__ == '__main__':
    logging.basicConfig()
    logging.getLogger().setLevel(logging.DEBUG)
    
    parser = argparse.ArgumentParser(description="OOProxy Arguments")
    parser.add_argument("--oo-host", metavar="OO_HOST", type=str, help="The OpenOffice Server Host", default="127.0.0.1")
    parser.add_argument("--oo-port", metavar="OO_PORT", type=int, help="The OpenOffice Server Port", default=8100)
    parser.add_argument("--port", metavar="PORT", type=int, help="The OOProxy Port", default=8099)
    parser.add_argument("--listen", metavar="LISTEN", type=str, help="Listen on Address", default="127.0.0.1")
    parser.add_argument("--timeout", metavar="TIMEOUT", type=int, help="Timeout in Seconds", default=30)
    parser.add_argument("--bufsize", metavar="BUFSIZE", type=int, help="Buffer size", default=32768)
    parser.add_argument("--syslog", dest='syslog', action="store_true", help="Log to syslog")
    args = parser.parse_args()
    
    if args.syslog:
        syslog_handler = logging.handlers.SysLogHandler(address="/dev/log")
        syslog_handler.setFormatter(logging.Formatter('%(name)s [%(levelname)s] %(message)s'))
        logging.getLogger().addHandler(syslog_handler)
    
    info("Server listen on port %s" % args.port)
    server = eventlet.listen((args.listen, args.port), backlog=50)
    pool = eventlet.GreenPool()
    while True:
        try:
            new_sock, address = server.accept()
            _logger.info("Client %s connected" % repr(new_sock.getpeername()))
            pool.spawn_n(application, new_sock.makefile("rwb"), new_sock, args)
        except (SystemExit, KeyboardInterrupt):
            break
    #from werkzeug.serving import run_simple
    #run_simple('localhost', DEFAUL_DEFAULT_PROXY_PORT, application)

