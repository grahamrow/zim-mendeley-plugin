# -*- coding: utf-8 -*-
#
# Copyright: 2014 Graham Rowlands <grahamrow@gmail.com>
# License: GNU GPL v2 or higher
#
# Uses API code from Mendeley's OpenOffice plugin to insert
# reference links that open in Mendeley. 

from __future__ import with_statement
from __future__ import division # We are doing math in this module ...

import logging
import re

from zim.plugins import PluginClass, extends, WindowExtension
from zim.errors import Error
from zim.actions import action
from zim.applications import Application, ApplicationError

import MendeleyDesktopAPI

logger = logging.getLogger('zim.plugins.insertcitation')

mendeleycmd = ('mendeleydesktop', '--help')

class MendeleyError(Error):
	description = _(
		'The Mendeley plugin was not able to fulfill\n'
		'this request. Please ensure that the Mendeley\n'
		'Desktop application is open.' )
		# T: error description

class DOIError(Error):
	description = _(
		'The Mendeley plugin was not able to fulfill\n'
		'this request. One or more DOI numbers may\n'
		'by undefined in your list of citations.' )
		# T: error description

class MendeleyPlugin(PluginClass):

	plugin_info = {
		'name': _('Mendeley Citations'), # T: plugin name
		'description': _('Mendeley is a free cross-platform desktop reference and paper management program (http://www.mendeley.com/).'
						'This plugin allows you to insert mendeley citations that link directly to the Mendeley desktop application '
						'or to a DOI URL. This is accomplished by interfacing with the Mendeley application, which must be open for the '
						'plugin to function. Any reference style from http://www.zotero.org/styles may be chosen.'), # T: plugin description
		'author': 'Graham Rowlands',
		'help': 'Plugins:Mendeley Citations',
	}

	global DOI, MENDELEY_LIB # Hack - to make sure translation is loaded
	MENDELEY_LIB = _('Mendeley Library') # T: option value
	DOI          = _('DOI Link') # T: option value

	@classmethod
	def check_dependencies(klass):
		has_mendeley = Application(mendeleycmd).tryexec()
		return has_mendeley, [('Mendeley Desktop', has_mendeley, True)]

	# Links either point to http://dx.doi.org/DOI-number-goes-here or mendeley://library/document/UUID-Number-goes-here
	# We let the use decide which to insert.
	plugin_preferences = (
		('citation_style', 'string', _('Citation Style'), 'apa'), # T: input label# key, type, label, default
		('bibliography_style', 'string', _('Bibliography Style'), 'physical-review-letters'), # T: input label# key, type, label, default
		('citation_link',  'choice', _('Links Points To'), MENDELEY_LIB, (MENDELEY_LIB, DOI)),
	)

	def get_uuid_from_link(self, url):
		if "dx.doi.org" in url:
			return url.split("=")[-1]
		else:
			return url.split("/")[-1]

	def insert_citation(self, buffer):
		style = self.preferences['citation_style'] or "apa"

		try:
			api = MendeleyDesktopAPI.MendeleyDesktopAPI("component context (unused)")
			api.resetCitations()
			api.setCitationStyle("http://www.zotero.org/styles/"+style)
			api.addCitationCluster(api.citation_choose_interactive(""))
			api.formatCitationsAndBibliography()

			uuids = api.getCitationClusterUUIDs(0)
			hrefs = []
			cites = []

			for uuid in uuids:
				citation = api.getFieldCodeFromUuid("{"+uuid+"}")
				api.resetCitations()
				api.addCitationCluster(citation)
				api.formatCitationsAndBibliography()
				if self.preferences['citation_link'] == MENDELEY_LIB:
					cites.append(api.getFormattedCitation(0))
					hrefs.append(api.getLocalURLs(0)[0])
					# buffer.insert_link_at_cursor(api.getFormattedCitation(0)+" ", href=api.getLocalURLs(0)[0])
				else:
					cites.append(api.getFormattedCitation(0))
					hrefs.append(api.getDOIURLs(0, addUUID=True)[0])
					# buffer.insert_link_at_cursor(api.getFormattedCitation(0)+" ", href=api.getDOIURLs(0)[0])
			
			# Defer insertion so that we either insert all (success) or none (failure)
			for cite, href, uuid in zip(cites, hrefs, uuids):
				buffer.insert_link_at_cursor(cite, href=href)
				buffer.insert_at_cursor(" ")

		except KeyError as error:
			msg = '%s: %s' % (error.__class__.__name__, error)
			raise DOIError, msg

		except Exception as error:
			msg = '%s: %s' % (error.__class__.__name__, error)
			raise MendeleyError, msg

	def html_to_zim(self, string):
		lines = string.splitlines()
		output = []
		ignores = [r'<!DOCTYPE',r'</*html', r'</*meta', r'</*title', r'</*head', r'</*body']
		replacements = [
				[r"</*?p.?>", "\n"],
				[r"<p *\w+='.*?'>", ""],
				[r"</*b.*?>", "**"],
				[r"</*i.*?>", "//"],
				[r"&nbsp;", ""],
				[r"</*span.*?>", ""],
				[r"</*.*?>", ""], # Catch All!
				[r" +", " "],
				]
		for line in lines:
			# Ignore lines that contains items in the above list
			if len([ignore for ignore in ignores if re.compile(ignore).search(line)!=None]) == 0:
				for old, new in replacements:
					line = re.sub(r'%s' % old, r'%s' % new, line)
				output.append(line)

		return "\n".join(output)

	def render_bibliography(self, uuids, buffer):
		api = MendeleyDesktopAPI.MendeleyDesktopAPI("component context (unused)")
		style = self.preferences['bibliography_style'] or "apa"
		api.setCitationStyle("http://www.zotero.org/styles/"+style)
		api.resetCitations()
		for uuid in uuids:
			citation = api.getFieldCodeFromUuid("{"+uuid+"}")
			api.addCitationCluster(citation)
		api.formatCitationsAndBibliography()
		link = api.getFormattedBibliography()
		with open(link, "r") as bibfile:
			buffer.insert_at_cursor(self.html_to_zim(bibfile.read()))

@extends('MainWindow')
class MainWindowExtension(WindowExtension):

	uimanager_xml = '''
	<ui>
	<menubar name='menubar'>
		<menu action='insert_menu'>
			<placeholder name='plugin_items'>
		 		<menuitem action='insert_citation'/>
		 	</placeholder>
		</menu>
		<menu action='tools_menu'>
			<placeholder name='plugin_items'>
				<menuitem action='generate_mendeley_bibliography'/>
			</placeholder>
		</menu>
	</menubar>
	</ui>
	'''

	@action(_('_Citation...'), '', '<Shift><Primary>C') # T: menu item
	def insert_citation(self):
		'''Action called by the menu item or key binding,
		will call the Mendeley API to insert a citation.
		'''
		buffer = self.window.ui.mainwindow.pageview.view.get_buffer()
		self.plugin.insert_citation(buffer)

	def get_mendeley_uuids(self):
		buffer = self.window.ui.mainwindow.pageview.view.get_buffer()
		uuids = []
		for link in self.window.ui.page.get_links():
			link_type, href, attrib = link 
			if 'mendeley://' in href or 'http://dx.doi.org' in href:
				uuids.append(self.plugin.get_uuid_from_link(href))
				buffer.insert_at_cursor("%s.\n" %href)
			# buffer.insert_at_cursor("Link of type %s and uuid %s.\n" % (link_type, self.plugin.get_uuid_from_link(href)))
		return uuids

	@action(_('_Generate Mendeley Bibliography'), '', '') # T: menu item
	def generate_mendeley_bibliography(self):
		buffer = self.window.ui.mainwindow.pageview.view.get_buffer()
		self.plugin.render_bibliography(self.get_mendeley_uuids(), buffer)
