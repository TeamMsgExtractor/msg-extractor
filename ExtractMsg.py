#!/usr/bin/env python -tt
# -*- coding: latin-1 -*-
"""
ExtractMsg:
	Extracts emails and attachments saved in Microsoft Outlook's .msg files

https://github.com/mattgwwalker/msg-extractor
"""

__author__ = "Matthew Walker, Edward Elliott"
__date__ = "2017-12-21"
__version__ = '0.4'

# --- LICENSE -----------------------------------------------------------------
#
#	Copyright 2013 Matthew Walker
#
#	This program is free software: you can redistribute it and/or modify
#	it under the terms of the GNU General Public License as published by
#	the Free Software Foundation, either version 3 of the License, or
#	(at your option) any later version.
#
#	This program is distributed in the hope that it will be useful,
#	but WITHOUT ANY WARRANTY; without even the implied warranty of
#	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#	GNU General Public License for more details.
#
#	You should have received a copy of the GNU General Public License
#	along with this program.  If not, see <http://www.gnu.org/licenses/>.

import os
import re
import sys
import random
import string
import StringIO
import optparse
import traceback
from email.parser import Parser as EmailParser
import email.utils
import olefile

SAVE_FMT = '%(sender)s - %(subject)s - %(desc)s'
ORIG_FMT = '%(date)s %(subject)s/$(desc)s'

# This property information was sourced from
# http://www.fileformat.info/format/outlookmsg/index.htm
# on 2013-07-22.
properties = {
	'001A': 'Message class',
	'0037': 'Subject',
	'003D': 'Subject prefix',
	'0040': 'Received by name',
	'0042': 'Sent repr name',
	'0044': 'Rcvd repr name',
	'004D': 'Org author name',
	'0050': 'Reply rcipnt names',
	'005A': 'Org sender name',
	'0064': 'Sent repr adrtype',
	'0065': 'Sent repr email',
	'0070': 'Topic',
	'0075': 'Rcvd by adrtype',
	'0076': 'Rcvd by email',
	'0077': 'Repr adrtype',
	'0078': 'Repr email',
	'007d': 'Message header',
	'0C1A': 'Sender name',
	'0C1E': 'Sender adr type',
	'0C1F': 'Sender email',
	'0E02': 'Display BCC',
	'0E03': 'Display CC',
	'0E04': 'Display To',
	'0E1D': 'Subject (normalized)',
	'0E28': 'Recvd account1 (uncertain)',
	'0E29': 'Recvd account2 (uncertain)',
	'1000': 'Message body',
	'1008': 'RTF sync body tag',
	'1035': 'Message ID (uncertain)',
	'1046': 'Sender email (uncertain)',
	'3001': 'Display name',
	'3002': 'Address type',
	'3003': 'Email address',
	'39FE': '7-bit email (uncertain)',
	'39FF': '7-bit display name',

	# Attachments (37xx)
	'3701': 'Attachment data',
	'3703': 'Attachment extension',
	'3704': 'Attachment short filename',
	'3707': 'Attachment long filename',
	'370E': 'Attachment mime tag',
	'3712': 'Attachment ID (uncertain)',

	# Address book (3Axx):
	'3A00': 'Account',
	'3A02': 'Callback phone no',
	'3A05': 'Generation',
	'3A06': 'Given name',
	'3A08': 'Business phone',
	'3A09': 'Home phone',
	'3A0A': 'Initials',
	'3A0B': 'Keyword',
	'3A0C': 'Language',
	'3A0D': 'Location',
	'3A11': 'Surname',
	'3A15': 'Postal address',
	'3A16': 'Company name',
	'3A17': 'Title',
	'3A18': 'Department',
	'3A19': 'Office location',
	'3A1A': 'Primary phone',
	'3A1B': 'Business phone 2',
	'3A1C': 'Mobile phone',
	'3A1D': 'Radio phone no',
	'3A1E': 'Car phone no',
	'3A1F': 'Other phone',
	'3A20': 'Transmit dispname',
	'3A21': 'Pager',
	'3A22': 'User certificate',
	'3A23': 'Primary Fax',
	'3A24': 'Business Fax',
	'3A25': 'Home Fax',
	'3A26': 'Country',
	'3A27': 'Locality',
	'3A28': 'State/Province',
	'3A29': 'Street address',
	'3A2A': 'Postal Code',
	'3A2B': 'Post Office Box',
	'3A2C': 'Telex',
	'3A2D': 'ISDN',
	'3A2E': 'Assistant phone',
	'3A2F': 'Home phone 2',
	'3A30': 'Assistant',
	'3A44': 'Middle name',
	'3A45': 'Dispname prefix',
	'3A46': 'Profession',
	'3A48': 'Spouse name',
	'3A4B': 'TTYTTD radio phone',
	'3A4C': 'FTP site',
	'3A4E': 'Manager name',
	'3A4F': 'Nickname',
	'3A51': 'Business homepage',
	'3A57': 'Company main phone',
	'3A58': 'Childrens names',
	'3A59': 'Home City',
	'3A5A': 'Home Country',
	'3A5B': 'Home Postal Code',
	'3A5C': 'Home State/Provnce',
	'3A5D': 'Home Street',
	'3A5F': 'Other adr City',
	'3A60': 'Other adr Country',
	'3A61': 'Other adr PostCode',
	'3A62': 'Other adr Province',
	'3A63': 'Other adr Street',
	'3A64': 'Other adr PO box',

	'3FF7': 'Server (uncertain)',
	'3FF8': 'Creator1 (uncertain)',
	'3FFA': 'Creator2 (uncertain)',
	'3FFC': 'To email (uncertain)',
	'403D': 'To adrtype (uncertain)',
	'403E': 'To email (uncertain)',
	'5FF6': 'To (uncertain)'}


def windowsUnicode (string):
	if string is None:
		return None
	if sys.version_info[0] >= 3:  # Python 3
		return str (string, 'utf_16_le')
	else:  # Python 2
		return unicode (string, 'utf_16_le')


def xstr (s):
	return '' if s is None else str (s)


def scrub_path (path):
	'''Create a valid safe path name.  Removes shell special chars.'''
	path = path.replace ('..', '')
	return re.sub (r'[^\w @/.,+_-]+', '-', path).strip ('-').strip ()


def scrub_name (name):
	'''Create a safe pathname component. Strips all slashes and special chars'''
	name = name.replace ('/', '')
	name = name.replace ('\\', '')
	return scrub_path (name)


def mkpath_uniq (path, always = False, format = '-%02d', start = 0):
	'''Make a unique path or die trying.
	always - append int?
	format - how to format appended num
	start - nonzero number to start uniqueness counter'''
	(base, ext) = os.path.splitext (path)

	num = start or 1

	if always:  uniq = '%s%s%s' % (base, (format % num), ext)
	else:	   uniq = path

	while os.path.exists (uniq):
		uniq = '%s%s%s' % (base, (format % num), ext)
		num += 1
	return uniq


class Attachment:
	def __init__ (self, msg, dir_):
		# Get long filename
		self.longFilename = msg._getStringStream ([dir_, '__substg1.0_3707'])

		# Get short filename
		self.shortFilename = msg._getStringStream ([dir_, '__substg1.0_3704'])

		# Get attachment data
		self.data = msg._getStream ([dir_, '__substg1.0_37010102'])
	

	@property
	def filename (self):
		return self.longFilename or self.shortFilename


	def save (self, save_path):
		# don't overwrite existing files
		save_path = mkpath_uniq (save_path)
		#verbose ('SAVING: %s' % save_path)

		f = open (save_path, 'wb')
		if self.data:
			f.write (self.data)
		else:
			print 'ERROR: no data found: %s' % save_path
		f.close ()
		return save_path


class Message (olefile.OleFileIO):
	def __init__ (self, filename):
		olefile.OleFileIO.__init__ (self, filename)

	def _getStream (self, filename):
		if self.exists (filename):
			stream = self.openstream (filename)
			return stream.read ()
		else:
			return None

	def _getStringStream (self, filename, prefer='unicode'):
		"""Gets a string representation of the requested filename.
		Checks for both ASCII and Unicode representations and returns
		a value if possible.  If there are both ASCII and Unicode
		versions, then the parameter /prefer/ specifies which will be
		returned.
		"""

		if isinstance (filename, list):
			# Join with slashes to make it easier to append the type
			filename = "/".join (filename)

		asciiVersion = self._getStream (filename + '001E')
		unicodeVersion = windowsUnicode (self._getStream (filename + '001F'))
		if asciiVersion is None:
			return unicodeVersion
		elif unicodeVersion is None:
			return asciiVersion
		else:
			if prefer == 'unicode':
				return unicodeVersion
			else:
				return asciiVersion

	@property
	def subject (self):
		return self._getStringStream ('__substg1.0_0037')

	@property
	def header (self):
		try:
			return self._header
		except Exception:
			headerText = self._getStringStream ('__substg1.0_007D')
			if headerText is not None:
				self._header = EmailParser ().parsestr (headerText)
			else:
				self._header = None
			return self._header

	@property
	def date (self):
		# Get the message's header and extract the date
		if self.header is None:
			return None
		else:
			return self.header['date']

	@property
	def parsedDate (self):
		return email.utils.parsedate (self.date)

	@property
	def sender_email (self):
		return self._getStringStream ('__substg1.0_0C1F')

	@property
	def sender_name (self):
		return self._getStringStream ('__substg1.0_0C1A')

	@property
	def sender (self):
		try:
			return self._sender
		except Exception:
			# Check header first
			if self.header is not None:
				headerResult = self.header["from"]
				if headerResult is not None:
					self._sender = headerResult
					return headerResult

			# Extract from other fields
			text = self._getStringStream ('__substg1.0_0C1A')
			email = self._getStringStream ('__substg1.0_0C1F')
			result = None
			if text is None:
				result = email
			else:
				result = text
				if email is not None:
					result = result + " <" + email + ">"

			self._sender = result
			return result

	@property
	def to (self):
		try:
			return self._to
		except Exception:
			# Check header first
			if self.header is not None:
				headerResult = self.header["to"]
				if headerResult is not None:
					self._to = headerResult
					return headerResult

			# Extract from other fields
			# TODO: This should really extract data from the recip folders,
			# but how do you know which is to/cc/bcc?
			display = self._getStringStream ('__substg1.0_0E04')
			self._to = display
			return display

	@property
	def cc (self):
		try:
			return self._cc
		except Exception:
			# Check header first
			if self.header is not None:
				headerResult = self.header["cc"]
				if headerResult is not None:
					self._cc = headerResult
					return headerResult

			# Extract from other fields
			# TODO: This should really extract data from the recip folders,
			# but how do you know which is to/cc/bcc?
			display = self._getStringStream ('__substg1.0_0E03')
			self._cc = display
			return display

	@property
	def body (self):
		# Get the message body
		return self._getStringStream ('__substg1.0_1000')

	@property
	def attachments (self):
		try:
			return self._attachments
		except Exception:
			# Get the attachments
			attachmentDirs = []

			for dir_ in self.listdir ():
				if dir_[0].startswith ('__attach') and dir_[0] not in attachmentDirs:
					attachmentDirs.append (dir_[0])

			self._attachments = []

			for attachmentDir in attachmentDirs:
				self._attachments.append (Attachment (self, attachmentDir))

			return self._attachments


	def get_text (self):
		'''Return the basic message headers and body as plain text.'''
		text = ''
		text += "From: %s\n" % xstr (self.sender)
		text += "To: %s\n" % xstr (self.to)
		text += "CC: %s\n" % xstr (self.cc)
		text += "Subject: %s\n" % xstr (self.subject)
		text += "Date: %s\n" % xstr (self.date)

		attachment_names = [x.filename for x in self.attachments] 
		text += "Attachments: %s\n" % '; '.join (attachment_names)
		text += "-----------------\n\n"
		text += self.body
		return text
	
	
	def get_json (self):
		'''Return the basic message headers and body as json formatted text.'''
		import json
		from imapclient.imapclient import decode_utf7

		attachment_names = [x.filename for x in self.attachments] 
		email_obj = {'from': xstr (self.sender),
					'to': xstr (self.to),
					'cc': xstr (self.cc),
					'subject': xstr (self.subject),
					'date': xstr (self.date),
					'attachments': attachmentNames,
					'body': decode_utf7 (self.body)}
		return json.dumps (email_obj, ensure_ascii = True)


	def get_save_path (self, save_fmt, infile, desc):
		'''Get the path to save output to, based on the requested format.'''
		msgfile = os.path.splitext (os.path.basename (infile)) [0]
		d = self.parsedDate
		date = d and '{0:02d}-{1:02d}-{2:02d}_{3:02d}{4:02d}'.format (*d)
		if not date: 
			date = 'unknown_date'
		# for security, scrub all data before exposing to filesystem
		subject = scrub_name (self.subject or '[no subject]')
		sender = scrub_name (self.sender_name or '[unknown sender]')
		email = scrub_name (self.sender_email or '[unknown sender_email]')
		desc = scrub_name (desc)

		save_path = SAVE_FMT % vars ()  # substitution magic
		save_path = scrub_path (save_path)
		save_path = mkpath_uniq (save_path)

		return save_path


	def save (self, infile, save_fmt = SAVE_FMT, json = False):
		'''Saves the message body and attachments found in the infile.
		Setting json to true will output the message body as JSON-formatted
		text.  The body and attachments are stored in a folder.  Setting
		useFileName to true will mean that the filename is used as the name of
		the folder; otherwise, the message's date and subject are used as the
		folder name.'''

		# make an output file name for the message body
		desc = 'message' + (json and '.json' or '.txt')
		save_path = self.get_save_path (save_fmt, infile, desc)

		# make sure dirs exist
		save_dir = os.path.dirname (save_path)
		if save_dir and not os.path.exists (save_dir):
			os.makedirs (save_dir)

		msg_text = json and self.get_json () or self.get_text ()

		f = open (save_path, 'wb')
		# fix unicode errors by specifying encoding
		f.write (msg_text.encode ('utf-8'))
		f.close ()

		# Save the attachments
		saved_files = []
		for att in self.attachments:
			desc = att.longFilename or att.shortFilename
			if not desc:
				salt = ''.join (random.choice (string.ascii_uppercase +
					string.digits) for _ in range (5))
				desc = 'unknown %s.bin' % salt

			att_path = self.get_save_path (save_fmt, infile, desc)
			saved_files.append (att.save (att_path))

		return save_path


	def save_raw (self, infile, save_fmt):
		'''Save the raw contents of the OLE file to a new directory. The
		directory name is created by appending -raw to save_path and ensuring
		the result is unique.'''
		save_path = self.get_save_path (save_fmt, infile, 'message')
		raw_dir = '%s-%s' % (save_path, 'raw')
		raw_dir = mkpath_uniq (raw_dir)
		os.makedirs (raw_dir)

		# Loop through all the directories in the OLE file
		for dir_ in self.listdir ():
			outfile = ''.join (dir_)
			code = dir_ [-1] [-8:-4]
			# shouldn't need global declaration, var isn't modified here
			#global properties
			if code in properties:
				outfile += ' - %s' % properties [code]

			# Generate appropriate filename
			outfile += '.dat'
			if dir_ [-1].endswith ("001E"):
				outfile += '.txt'

			save_path = os.path.join (raw_dir, outfile)

			# Save contents of directory
			f = open (save_path, 'wb')
			f.write (self._getStream (dir_))
			f.close ()


	def dump (self):
		# Prints out a summary of the message
		print 'Message'
		print 'Subject:', self.subject
		print 'Date:', self.date
		print 'Body:'
		print self.body


	def debug (self):
		for dir_ in self.listdir ():
			if dir_[-1].endswith ('001E'):  # FIXME: Check for unicode 001F too
				print ("Directory: " + str (dir))
				print ("Contents: " + self._getStream (dir))


# Appears to be dead code.  This function isn't called anywhere.  Should be
# removed.
# If it's somehow needed (say, called by OleFile) then need to add save_path
# arg to save () call.
#
#	def save_attachments (self, raw=False):
#		"""Saves only attachments in the same folder.
#		"""
#		for attachment in self.attachments:
#			attachment.save ()


def main (args):
	pass

def main (args):
	usage = 'usage: %s [options] <file> [file2 ...]\n' % sys.argv [0]
	desc = """Parses Microsoft Outlook Message (.msg) files and saves their contents to the
current directory.  On error the script will write out a 'raw' directory will
all the details from the file, but in a less-than-desirable format. To force
this mode, the flag '--raw' can be specified.

For more control, the save file location can be set with the --save_format
flag.  You can include these fields from the message in your output location:
	sender - sender's name
	email - sender's email
	subject - message subject
	date - message sent date
	msgfile - the filename of the .msg file being read
	desc - data description. for attachments it's the embedded attachment file
		name.  for the message body it's 'message.txt'
List the desired field as '%(field)s' and it will substituted.  Use / if you
want to use subdirectories.

Examples
	in current directory as date and subject
		-s '%(date)s %(subject)s'

	in current directory as sender, subject, and date separatd by dashes
		-s '%(sender)s - %(subject)s - %(date)s'

	under a subdir named sender in another subdir named subject with files
	named by the desc
		-s '%(sender)s/%(subject)s/%(desc)s'
"""

	parser = optparse.OptionParser ('\n'.join ((usage, __doc__, desc)))

	parser.add_option ('-j', '--json', action = 'store_true',
		help = 'output message in json format')

	parser.add_option ('-r', '--raw', action = 'store_true', default = False,
		help = 'output message in raw format [default]')

	parser.add_option ('-m', '--use_msg_file', action = 'store_true',
		help = 'use only the input .msg filename as output directory name. same as: -s "%(msgfile)s/". overrides -s flag.')

	parser.add_option ('-n', '--nowork', action = 'store_true',
		help = 'dont create any files, just show what would be done')

	parser.add_option ('-o', '--orig_format', action = 'store_true',
		help = 'use the original output format: -s "%(date)s %(subject)s/%(desc)s". overrides -s flag.')

	parser.add_option ('-s', '--save_format', action = 'store', default = SAVE_FMT,
		help = 'format string for outputting saved message file name')

	parser.add_option ('-q', '--quiet', action = 'count', help = 'less verbose')
	parser.add_option ('-v', '--verbose', action = 'count', help = 'more verbose')

	opts, args = parser.parse_args (args)

	# move to options section
	if opts.use_msg_file:
		opts.save_fmt = '%(msgfile)s/'

	if opts.orig_format:
		opts.save_fmt = '%(date)s %(subject)s/%(desc)s'

	if opts.nowork:
		# replace functions that modify the filesystem with dummy versions
		def echo_func (prefix, ret_fn):
			def echo (*args):
				sys.stdout.write ('%s: %s\n' % (prefix, ' '.join (args)))
				if ret_fn: return ret_fn ()
			return echo

		os.makedirs = echo_func ('make dirs', None)
		global open
		open = echo_func ('open file', StringIO.StringIO)


	for arg in args:
		msg = Message (arg)
		if opts.raw:
			msg.save_raw (arg, opts.save_format)
		else:
			try: 
				msg.save (arg, opts.save_format, opts.json)
			except:
				msg.save_raw (arg, opts.save_format)


if __name__ == "__main__":
	main (sys.argv [1:])

