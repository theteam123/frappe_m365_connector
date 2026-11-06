# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""Unit tests for working out who a Send to ERP email is really from.

Paul, 2026-11-07: forwarded emails come "from" a colleague. The original sender is read from
the forwarded-message header, our own staff are skipped, and a picked Contact is checked. The
database is patched, so these run without records. Run them directly:

    cd sites && ../env/bin/python -m unittest m365email.m365email.test_send_to_erp_people
"""

import time
import unittest
from unittest.mock import patch

import frappe

from m365email.m365email import send_to_erp_people as people

MAX_CRAFTED_BODY_SECONDS = 1
STAFF = {"angela.obrien@sgcaustralia.com.au", "rik.bidgood@sgcaustralia.com.au"}

# NOTE: The three header layouts seen on real forwarded emails on staging, 2026-11-07.
OUTLOOK_DESKTOP_FORWARD = (
	'<div><b>From:</b> Bob Tester &lt;<a href="mailto:bob@miw.net.au">bob@miw.net.au</a>&gt;<br>'
	"<b>Sent:</b> Friday, 6 November 2026 11:07 AM<br><b>To:</b> Angela O'Brien "
	"&lt;angela.obrien@sgcaustralia.com.au&gt;<br><b>Subject:</b> Quote MBL102897</div>"
)
WORD_STYLE_FORWARD = (
	'<p class="MsoNormal"><b><span lang="EN-US">From:</span></b><span lang="EN-US"> '
	"Rik Bidgood &lt;rik.bidgood@sgcaustralia.com.au&gt; <br><b>Sent:</b> Friday</span></p>"
	"<p><b>From:</b> Sales &lt;sales@miw.net.au&gt;</p>"
)
BARE_ADDRESS_FORWARD = "<p>-----Original Message-----<br>From: quotes@miw.net.au<br>Sent: today</p>"



def FakeExists(doctype: str, filters: object) -> bool:
	if doctype != "User" or not isinstance(filters, dict):
		return False
	if filters.get("user_type") != "System User":
		return False
	return (filters.get("email") or filters.get("name")) in STAFF



def RaiseValidationError(message: str, *args: object, **kwargs: object) -> None:
	raise frappe.ValidationError(message)



@patch.object(people.frappe.db, "exists", side_effect=FakeExists)
class TestReadForwardedSenders(unittest.TestCase):
	def test_reads_outlook_desktop_header(self, *mocks: object) -> None:
		senders = people.ReadForwardedSenders(OUTLOOK_DESKTOP_FORWARD)
		self.assertEqual(senders, [{"name": "Bob Tester", "email": "bob@miw.net.au"}])


	def test_reads_every_header_newest_first(self, *mocks: object) -> None:
		emails = [sender["email"] for sender in people.ReadForwardedSenders(WORD_STYLE_FORWARD)]
		self.assertEqual(emails, ["rik.bidgood@sgcaustralia.com.au", "sales@miw.net.au"])


	def test_reads_bare_address(self, *mocks: object) -> None:
		senders = people.ReadForwardedSenders(BARE_ADDRESS_FORWARD)
		self.assertEqual(senders, [{"name": "", "email": "quotes@miw.net.au"}])


	def test_plain_email_has_no_forward_header(self, *mocks: object) -> None:
		self.assertEqual(people.ReadForwardedSenders("<p>Please find our quote attached.</p>"), [])


	def test_crafted_body_reads_quickly(self, *mocks: object) -> None:
		crafted = "From: " + "a" * 40000 + "\nFrom:" + " " * 40000
		started = time.monotonic()
		people.ReadForwardedSenders(crafted)
		self.assertLess(time.monotonic() - started, MAX_CRAFTED_BODY_SECONDS)



@patch.object(people.frappe.db, "exists", side_effect=FakeExists)
class TestGuessPerson(unittest.TestCase):
	def test_external_sender_is_the_person(self, *mocks: object) -> None:
		person = people.GuessPerson("Bob Tester", "Bob@MIW.net.au", OUTLOOK_DESKTOP_FORWARD)
		self.assertEqual(person["email"], "bob@miw.net.au")
		self.assertFalse(person["is_forwarded"])


	def test_colleague_forward_names_original_sender(self, *mocks: object) -> None:
		email = "angela.obrien@sgcaustralia.com.au"
		person = people.GuessPerson("Angela O'Brien", email, OUTLOOK_DESKTOP_FORWARD)
		self.assertEqual((person["name"], person["email"]), ("Bob Tester", "bob@miw.net.au"))
		self.assertTrue(person["is_forwarded"])


	def test_colleagues_in_the_chain_are_skipped(self, *mocks: object) -> None:
		email = "angela.obrien@sgcaustralia.com.au"
		person = people.GuessPerson("Angela O'Brien", email, WORD_STYLE_FORWARD)
		self.assertEqual(person["email"], "sales@miw.net.au")


	def test_only_colleagues_is_flagged_internal(self, *mocks: object) -> None:
		email = "angela.obrien@sgcaustralia.com.au"
		person = people.GuessPerson("Angela O'Brien", email, "<p>No forward here</p>")
		self.assertTrue(person["is_internal"])
		self.assertFalse(person["is_forwarded"])



class TestResolveChosenContact(unittest.TestCase):
	def setUp(self) -> None:
		patcher = patch.object(people.frappe, "throw", side_effect=RaiseValidationError)
		patcher.start()
		self.addCleanup(patcher.stop)


	def test_none_picked_is_none(self) -> None:
		self.assertIsNone(people.ResolveChosenContact(""))


	def test_missing_contact_is_refused(self) -> None:
		with patch.object(people.frappe.db, "exists", return_value=False):
			with self.assertRaises(frappe.ValidationError):
				people.ResolveChosenContact("no-such-contact")


	def test_unreadable_contact_is_refused(self) -> None:
		with patch.object(people.frappe.db, "exists", return_value=True):
			with patch.object(people.frappe, "has_permission", return_value=False):
				with self.assertRaises(frappe.ValidationError):
					people.ResolveChosenContact("someone-else")


	def test_readable_contact_is_kept(self) -> None:
		with patch.object(people.frappe.db, "exists", return_value=True):
			with patch.object(people.frappe, "has_permission", return_value=True):
				self.assertEqual(people.ResolveChosenContact("bob"), "bob")



	def test_missing_and_unreadable_give_the_same_message(self) -> None:
		messages = []
		for exists, readable in ((False, True), (True, False)):
			with patch.object(people.frappe.db, "exists", return_value=exists):
				with patch.object(people.frappe, "has_permission", return_value=readable):
					with self.assertRaises(frappe.ValidationError) as caught:
						people.ResolveChosenContact("x")
			messages.append(str(caught.exception))
		self.assertEqual(messages[0], messages[1])



class TestCraftedHTML(unittest.TestCase):
	def test_many_open_brackets_read_quickly(self) -> None:
		started = time.monotonic()
		people.ReadForwardedSenders("<" * 2_000_000)
		self.assertLess(time.monotonic() - started, MAX_CRAFTED_BODY_SECONDS)



	def test_long_attribute_is_not_read_as_a_header(self) -> None:
		body = '<div title="' + "x" * 3000 + '\nFrom: fake@example.com\n">Visible text</div>'
		with patch.object(people.frappe.db, "exists", side_effect=FakeExists):
			self.assertEqual(people.ReadForwardedSenders(body), [])



	def test_long_line_break_tag_still_breaks_the_line(self) -> None:
		body = 'intro<br title="' + "a" * 300 + '">From: real@example.com'
		with patch.object(people.frappe.db, "exists", side_effect=FakeExists):
			senders = people.ReadForwardedSenders(body)
		self.assertEqual([sender["email"] for sender in senders], ["real@example.com"])
