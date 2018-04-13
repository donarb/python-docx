# encoding: utf-8

"""
The :mod:`pptx.packaging` module coheres around the concerns of reading and
writing presentations to and from a .pptx file.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)


class AppProperties(object):
    """
    Corresponds to part named ``/docProps/app.xml``, containing the core
    document properties for this document package.
    """
    def __init__(self, element):
        self._element = element

    @property
    def template(self):
        return self._element.template_text

    @property
    def totaltime(self):
        return self._element.totaltime_number

    @property
    def pages(self):
        return self._element.pages_number

    @property
    def words(self):
        return self._element.words_number

    @property
    def characters(self):
        return self._element.characters_number

    @property
    def application(self):
        return self._element.application_text

    @property
    def docsecurity(self):
        return self._element.docsecurity_number

    @property
    def lines(self):
        return self._element.lines_number

    @property
    def paragraphs(self):
        return self._element.paragraphs_number

    @property
    def company(self):
        return self._element.company_text

    @property
    def appversion(self):
        return self._element.appversion_number
