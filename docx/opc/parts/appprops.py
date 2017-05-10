# encoding: utf-8

"""
App properties part, corresponds to ``/docProps/app.xml`` part in package.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)


from ..constants import CONTENT_TYPE as CT
from ..appprops import AppProperties
from ...oxml.appprops import CT_AppProperties
from ..packuri import PackURI
from ..part import XmlPart


class AppPropertiesPart(XmlPart):
    """
    Corresponds to part named ``/docProps/app.xml``, containing the app
    document properties for this document package.
    """
    @classmethod
    def default(cls, package):
        """
        Return a new |AppPropertiesPart| object initialized with default
        values for its base properties.
        """
        app_properties_part = cls._new(package)
        return app_properties_part

    @property
    def app_properties(self):
        """
        A |AppProperties| object providing read/write access to the app
        properties contained in this app properties part.
        """
        return AppProperties(self.element)

    @classmethod
    def _new(cls, package):
        partname = PackURI('/docProps/app.xml')
        content_type = CT.OFC_EXTENDED_PROPERTIES
        appProperties = CT_AppProperties.new()
        return AppPropertiesPart(
            partname, content_type, appProperties, package
        )
