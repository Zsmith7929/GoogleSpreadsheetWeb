#!/usr/bin/python
#
# Copyright (C) 2015 Zacheriah Smith
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#	   http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

__author__ = 'Smith7929@gmail.com (Zacheriah Smith)'

try:
	from xml.etree import ElementTree
except ImportError:
	from elementtree import ElementTree
import sys
import gdata.spreadsheet.service
import gdata.service
import atom.service
import gdata.spreadsheet
import atom
import getopt
import string
import cgi
import json
import urllib
import logging
import re

from bs4 import BeautifulSoup

import webapp2

COLUMN_MAP = {"a":1,"b":2,"c":3,"d":4,"e":5,"f":6,"g":7,"h":8,"i":9,"j":10,
			"k":11,"l":12,"m":13,"n":14,"o":15,"p":16,"q":17,"r":18,"s":19,
			"t":20,"u":21,"v":22,"w":23,"x":24,"y":25,"z":26, "aa":27,"ab":28,"ac":29}

class TableMagic(object):

	def __init__(self):
		self.sheetDict = {}
	
	def _returnCell(self, key):
	
		linkCells = ["atwill","encounter","daily","utility","equipment","rituals"]
		
		nameShort = json.dumps(self.sheetDict[key][1][:10]).replace('"', '').replace("'","")
		name = json.dumps(self.sheetDict[key][1]).replace('"', '').replace("'","")
		category = self.sheetDict[key][0].capitalize()
		
	
		if self.sheetDict[key][0] in linkCells:
			x = "<label for=" + key + ">" + category + ":</label><button onclick='search(this.innerHTML)' class='searchButton'>"+name+"</button>"
		
		else:
		
			x = "<label for=" + key + ">" + category + ":</label> <input type='text' onkeypress='return update(event)' size='6' name='" + key + "' id='" + Creds.curr_key + "' value='" + name + "'>"
		return x
	
	def getSheetTable(self, spreadsheetId=None):

		self.sheetDict = {}

		if spreadsheetId != None:
			Creds.curr_key = spreadsheetId
			feed = Creds.gd_client.GetWorksheetsFeed(Creds.curr_key)
			id_parts = feed.entry[0].id.text.split('/')
			Creds.curr_wksht_id = id_parts[len(id_parts) - 1]
			cellFeed = Creds.gd_client.GetCellsFeed(Creds.curr_key, Creds.curr_wksht_id)

			for entry in cellFeed.entry:
				if entry.title.text.lower() in Creds.cellDict:
					self.sheetDict.setdefault(entry.title.text.lower(), [Creds.cellDict[entry.title.text.lower()], str(entry.content.text)])
					
					
		for key in Creds.cellDict:
			if key not in self.sheetDict:
				self.sheetDict.setdefault(key, [Creds.cellDict[key],"n/a"])
					
		with open("table.html") as f:
			table = f.read().format(self._returnCell("w1"), self._returnCell("a2"), self._returnCell("l2"), self._returnCell("m4"),
								self._returnCell("k2"), self._returnCell("z2"), self._returnCell("a4"), self._returnCell("k20"),
								self._returnCell("a22"), self._returnCell("b25"), self._returnCell("h22"), self._returnCell("h25"),
								self._returnCell("h27"), self._returnCell("d30"), self._returnCell("a12"), self._returnCell("a13"),
								self._returnCell("a14"), self._returnCell("a15"), self._returnCell("a16"), self._returnCell("a17"),
								self._returnCell("k8"), self._returnCell("k12"), self._returnCell("k14"), self._returnCell("k16"),
								self._returnCell("u8"), self._returnCell("u12"), self._returnCell("u13"), self._returnCell("d29"),
								self._returnCell("a34"), self._returnCell("a35"), self._returnCell("a36"), self._returnCell("a37"),
								self._returnCell("a38"), self._returnCell("a39"), self._returnCell("a40"), self._returnCell("a41"),
								self._returnCell("a42"), self._returnCell("a43"), self._returnCell("a44"), self._returnCell("a45"),
								self._returnCell("a46"), self._returnCell("a47"), self._returnCell("a48"), self._returnCell("a49"),
								self._returnCell("a50"), self._returnCell("a97"), self._returnCell("a98"),self._returnCell("a99"),
								self._returnCell("u42"), self._returnCell("u43"), self._returnCell("u44"), self._returnCell("u45"),
								self._returnCell("a54"), self._returnCell("a55"), self._returnCell("a56"), self._returnCell("a57"),
								self._returnCell("a61"), self._returnCell("a62"), self._returnCell("a63"),self._returnCell("a64"),
								self._returnCell("a68"), self._returnCell("a69"), self._returnCell("a70"),self._returnCell("a71"),
								self._returnCell("a75"), self._returnCell("a76"), self._returnCell("a77"), self._returnCell("a78"),
								self._returnCell("a85"), self._returnCell("a86"), self._returnCell("a87"), self._returnCell("a88"),
								self._returnCell("k85"), self._returnCell("k86"), self._returnCell("k87"), self._returnCell("k88"),
								spreadsheetId)
				
		return table

class SpreadsheetServices(object):
	
	def __init__(self, email=None, password=None):
		self.gd_client = gdata.spreadsheet.service.SpreadsheetsService()
		self.gd_client.email = email
		self.gd_client.password = password
		self.gd_client.source = 'Webapp'
		self.curr_key = ''
		self.curr_wksht_id = ''
		self.cellDict = {}
		
	def createCellDict(self):
		f = open("cellDict.txt")
		self.cellDict = json.load(f)
	
	def getSpreadsheetListBox(self):
		feed = self.gd_client.GetSpreadsheetsFeed()
		spreadsheetDict = {}
		for entry in feed.entry:
			spreadsheetDict.setdefault(entry.title.text, entry.id.text.split("/")[len(entry.id.text.split("/"))-1])		
		return spreadsheetDict
					
	def login(self):
		self.gd_client.ProgrammaticLogin()
		
class SheetHandler(webapp2.RequestHandler):
		
	def get(self):
		self.response.write("You aint 'posed to GET this shiz")
		
	def post(self):
		try:
			spreadsheetId = cgi.escape(self.request.get("one"))
			Table = TableMagic()
			self.response.out.headers['Content-Type'] = 'text/html'
			self.response.out.write(Table.getSheetTable(spreadsheetId))
		except:
			self.redirect("/")

class QueryHandler(webapp2.RequestHandler):
		
	def get(self):
		self.response.write("You aint 'posed to GET this shiz")
		
	def post(self):
		query = re.sub(r'\([^)]*\)', '', cgi.escape(self.request.get("query")))
		logging.info("QUERY: $$$$$$$$$$$: "+query)
		page = urllib.urlopen("http://dnd4.wikia.com/wiki/Special:Search?search="+query+"&fulltext=Search")
		
		soup = BeautifulSoup(page)
		tag = soup.find("a", {"data-pos" : "1"})
		link = tag.get("href")
		
		self.response.out.headers['Content-Type'] = 'text/html'
		self.response.write(link)

class UpdateHandler(webapp2.RequestHandler):

	def get(self):
		self.response.write("NO.")
		
	def post(self):
		Creds.login() # RELOG IN CASE OF TIMEOUT
		cell = cgi.escape(self.request.get("cell"))
		spreadsheetId = cgi.escape(self.request.get("spreadsheet"))
		updatedValue = cgi.escape(self.request.get("updatedValue"))
		cellCol = COLUMN_MAP[''.join([str(x) for x in cell if x in string.letters])]
		cellRow = ''.join([x for x in cell if x in string.digits])
		
		feed = Creds.gd_client.GetWorksheetsFeed(Creds.curr_key)
		id_parts = feed.entry[0].id.text.split('/')
		wrkSheetId = id_parts[len(id_parts) - 1]
		entry = Creds.gd_client.UpdateCell(row=int(cellRow), col=int(cellCol), inputValue=updatedValue, key=spreadsheetId, wksht_id=wrkSheetId)
		
		newTableObj = TableMagic()
		newTable = newTableObj.getSheetTable(spreadsheetId)
		
		self.response.write([spreadsheetId, newTable])
		
	

class MainHandler(webapp2.RequestHandler):

	def handle_exception(self, exception, debug):
	
		logging.exception(exception)

		self.response.write('<h1 style="text-align:center">An error occurred: '+str(exception)+'</h1>')

		if isinstance(exception, webapp2.HTTPException):
			self.response.set_status(exception.code)
		else:
			self.response.set_status(500)

			
	def writePage(self, credentials=False, server=None):
	
		self.response.write("""
		
		<!DOCTYPE html>
		<html>
		<head>
		<meta charset="UTF-8">
		<title>D&D 4ed WebApp</title>
		<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
		<script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/jquery-ui.min.js"></script>
		<script> 
		
		function update(e) {
			if (e.keyCode === 13) {
				$(e.target).parents("table").first().find("input,button,textarea").attr("disabled","disabled");
				e.target.style.backgroundColor="red";
				$(e.target).animate({backgroundColor: '#FFFFFF'}, 1800);
				$.ajax({
					url: '/update',
					type: 'POST',
					data: {"cell" : e.target.name,
							"spreadsheet" : e.target.id,
							"updatedValue" : e.target.value},
					success: function(data) {
						var p = $(e.target).parents('div').first().find('.load').first();
						p.trigger("click");
						
						$(e.target).parents("table").first().find("input,button,textarea").removeAttribute("disabled");
						
					}
				});
			}
		}
		
		function search(url) {
			$.ajax({
				url: '/query',
				type: 'POST',
				data: {"query" : url},
				success: function(data) {
					var myTab = window.open("about:blank","myPopup");
					myTab.location = data;
				}
			});
			
		}
		function handleClick(e){
			$('#'+e.target.id).attr('disabled', 'disabled');
			var value = $('#spreadsheet'+e.target.id).val();
			$('#gif'+e.target.id).css("visibility","visible");
			$.ajax({
				url: '/sheet',
				type: 'POST',
				data: {"one" : value},
				success: function(data) {
					$("#div"+e.target.id).find("table").html(data);
					$('#gif'+e.target.id).css("visibility","hidden");
					$('#'+e.target.id).removeAttr('disabled');
				}
			});
		};
		$(document).ready(function(){
			$('.load').click(handleClick);
			
		});
		</script>
		
		<style>""")
		
		if not credentials:
			self.response.write("""
				body {
				background-image : url("http://i.imgur.com/OksyBuG.jpg");
				background-size: cover;
				}""")

		self.response.write("""

h1, table { text-align: center; }

table { border:1px solid #000000; border-collapse: collapse;  width: 25%; margin: 0 auto 5rem;}

.searchButton { width:8em; }

th, td { white-space: nowrap; font-size: 1.0rem; }

tr {background: hsl(240, 80%, 80%); }

tr, td { transition: .4s ease-in; } 

tr:nth-child(even) { background: hsla(240, 80%, 80%, 0.7); }

td:empty {background: hsla(50, 25%, 60%, 0.7); }

img { margin-left: 10px; visibility:hidden; }

.divLeft { text-align:center; width: 500; height:500; margin-left:auto; margin-right:auto; float:left}
.divRight { text-align:center; width: 500; height:500; margin-left:auto; margin-right:auto%; float:right}
.rowDiv { clear: both; }
.signIn { text-align:right; margin-left:auto; margin-right:225px; float:right; }

		</style>
		
		</head>
		<body>
		""")

		if not credentials:
			self.response.write("""
			<br><br><br><br><br><br><br><br><br><br>
			<div class="signIn">
			<h1 style="color:#FFCC66;">Login:
			<form action="/" method="post">
			<input type="text" name="user" value="gmailAccount@gmail.com">
			<input type="password" name="password">
			<input type="submit">
			</form>
			</div>
			""")
		else:
			Table = TableMagic()
			spreadsheetDict = Creds.getSpreadsheetListBox()
			self.response.write("""
			<div id="rowOne" class="rowDiv"> 
			<div id="divOne" class="divLeft">""")
			self.response.write("<select id='spreadsheetOne'>")
			for key in spreadsheetDict:
				self.response.write("<option value="+spreadsheetDict[key]+">"+key+"</option>")
			self.response.write("""
			</select>
			<input type='button' class='load' id='One' value='Load'><img class="loading" src="http://i.stack.imgur.com/MnyxU.gif" id="gifOne" height="20" width="20">
			""")
			self.response.write(Table.getSheetTable())
			self.response.write("""
			</div>
			<div id="divTwo" class="divRight">""")
			self.response.write("<select id='spreadsheetTwo'>")
			for key in spreadsheetDict:
				self.response.write("<option value="+spreadsheetDict[key]+">"+key+"</option>")
			self.response.write("""
			</select>
			<input type='button' class ='load' id='Two' value='Load'><img class="loading" src="http://i.stack.imgur.com/MnyxU.gif" id="gifTwo" height="20" width="20">""")
			self.response.write(Table.getSheetTable())
			self.response.write("""
			</div>
			</div>""")
			
			self.response.write("""
			<br><br><br><br><br><br><br><br>
			<div id="rowTwo" class="rowDiv"> 
			<div id=divThree class="divLeft">""")
			
			self.response.write("<select id='spreadsheetThree'>")
			for key in spreadsheetDict:
				self.response.write("<option value="+spreadsheetDict[key]+">"+key+"</option>")
			self.response.write("""
			</select>
			<input type='button' class ='load' id='Three' value='Load'><img class="loading" src="http://i.stack.imgur.com/MnyxU.gif" id="gifThree" height="20" width="20">""")
			self.response.write(Table.getSheetTable())
			self.response.write("""
			</div>
			<div id="divFour" class="divRight">""")
			self.response.write("<select id='spreadsheetFour'>")
			for key in spreadsheetDict:
				self.response.write("<option value="+spreadsheetDict[key]+">"+key+"</option>")
			self.response.write("""
			</select>
			<input type='button' class ='load' id='Four' value='Load'><img class="loading" src="http://i.stack.imgur.com/MnyxU.gif" id="gifFour" height="20" width="20">""")
			self.response.write(Table.getSheetTable())
			self.response.write("""
			</div>
			</div>""")
			
			
		self.response.write("</body></html>")
	
	
	def get(self):
		self.writePage()
		
	def post(self):
		email = cgi.escape(self.request.get("user"))
		password = cgi.escape(self.request.get("password"))
		Creds.gd_client.email = email
		Creds.gd_client.password = password
		Creds.login()
		Creds.createCellDict()
		if Creds.gd_client.email is not None:
			self.writePage(credentials=True,server=Creds)
		else:
			self.writePage()
	


#------------------------------------------------------#
Creds = SpreadsheetServices()
app = webapp2.WSGIApplication([																																										 
	('/', MainHandler),
	('/sheet', SheetHandler),
	('/query', QueryHandler),
	('/update', UpdateHandler)
], debug=True)
