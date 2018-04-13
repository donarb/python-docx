# encoding: utf-8

"""
lxml custom element classes for app properties-related XML elements.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import re

from datetime import datetime, timedelta

from . import parse_xml
from .xmlchemy import BaseOxmlElement, ZeroOrOne


class CT_AppProperties(BaseOxmlElement):
    """
    ``<cp:AppProperties>`` element, the root element of the App Properties
    part stored as ``/docProps/app.xml``. Implements many of the Dublin Core
    document metadata elements. String elements resolve to an empty string
    ('') if the element is not present in the XML. String elements are
    limited in length to 255 unicode characters.
    """
    template = ZeroOrOne('template', successors=())
    totaltime = ZeroOrOne('totaltim', successors=())
    pages = ZeroOrOne('pages', successors=())
    words = ZeroOrOne('words', successors=())
    characters = ZeroOrOne('characters', successors=())
    application = ZeroOrOne('application', successors=())
    docsecurity = ZeroOrOne('docsecurity', successors=())
    lines = ZeroOrOne('lines', successors=())
    paragraphs = ZeroOrOne('paragraphs', successors=())
    company = ZeroOrOne('company', successors=())
    appversion = ZeroOrOne('appversion', successors=())

    _appProperties_tmpl =  '<cp:appProperties />\n'


    @classmethod
    def new(cls):
        """
        Return a new ``<appProperties>`` element
        """
        xml = cls._appProperties_tmpl
        appProperties = parse_xml(xml)
        return appProperties


    @property
    def template(self):
        return self._text_of_element('template')

    @property
    def totaltime(self):
        return self._text_of_element('totaltime')

    @property
    def pages(self):
        return self._number_of_element('pages')

    @property
    def words(self):
        return self._number_of_element('words')

    @property
    def characters(self):
        return self._number_of_element('characters')

    @property
    def application(self):
        return self._number_of_text_of_element('application')

    @property
    def docsecurity(self):
        return self._number_of_element('docsecurity')

    @property
    def lines(self):
        return self._number_of_element('lines')

    @property
    def paragraphs(self):
        return self._number_of_element('paragraphs')

    @property
    def company(self):
        return self._text_of_element('company')

    @property
    def appversion(self):
        return self._number_of_element('appversion')

    def _datetime_of_element(self, property_name):
        element = getattr(self, property_name)
        if element is None:
            return None
        datetime_str = element.text
        try:
            return self._parse_W3CDTF_to_datetime(datetime_str)
        except ValueError:
            # invalid datetime strings are ignored
            return None

    def _get_or_add(self, prop_name):
        """
        Return element returned by 'get_or_add_' method for *prop_name*.
        """
        get_or_add_method_name = 'get_or_add_%s' % prop_name
        get_or_add_method = getattr(self, get_or_add_method_name)
        element = get_or_add_method()
        return element

    @classmethod
    def _offset_dt(cls, dt, offset_str):
        """
        Return a |datetime| instance that is offset from datetime *dt* by
        the timezone offset specified in *offset_str*, a string like
        ``'-07:00'``.
        """
        match = cls._offset_pattern.match(offset_str)
        if match is None:
            raise ValueError(
                "'%s' is not a valid offset string" % offset_str
            )
        sign, hours_str, minutes_str = match.groups()
        sign_factor = -1 if sign == '+' else 1
        hours = int(hours_str) * sign_factor
        minutes = int(minutes_str) * sign_factor
        td = timedelta(hours=hours, minutes=minutes)
        return dt + td

    _offset_pattern = re.compile('([+-])(\d\d):(\d\d)')

    @classmethod
    def _parse_W3CDTF_to_datetime(cls, w3cdtf_str):
        # valid W3CDTF date cases:
        # yyyy e.g. '2003'
        # yyyy-mm e.g. '2003-12'
        # yyyy-mm-dd e.g. '2003-12-31'
        # UTC timezone e.g. '2003-12-31T10:14:55Z'
        # numeric timezone e.g. '2003-12-31T10:14:55-08:00'
        templates = (
            '%Y-%m-%dT%H:%M:%S',
            '%Y-%m-%d',
            '%Y-%m',
            '%Y',
        )
        # strptime isn't smart enough to parse literal timezone offsets like
        # '-07:30', so we have to do it ourselves
        parseable_part = w3cdtf_str[:19]
        offset_str = w3cdtf_str[19:]
        dt = None
        for tmpl in templates:
            try:
                dt = datetime.strptime(parseable_part, tmpl)
            except ValueError:
                continue
        if dt is None:
            tmpl = "could not parse W3CDTF datetime string '%s'"
            raise ValueError(tmpl % w3cdtf_str)
        if len(offset_str) == 6:
            return cls._offset_dt(dt, offset_str)
        return dt

    def _text_of_element(self, property_name):
        """
        Return the text in the element matching *property_name*, or an empty
        string if the element is not present or contains no text.
        """
        element = getattr(self, property_name)
        if element is None:
            return ''
        if element.text is None:
            return ''
        return element.text

    def _number_of_element(self, property_name):
        """
        Return the number in the element matching *property_name*, or zero
        if the element is not present or contains no value.
        """
        element = getattr(self, property_name)
        if element is None:
            return 0
        element_str = element.text
        try:
            element = int(element_str)
        except ValueError:
            # non-integer  strings also resolve to 0
            element = 0
        # as do negative integers
        if element < 0:
            element = 0
        return element
