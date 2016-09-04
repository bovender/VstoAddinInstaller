{
=====================================================================
== inc\constants.iss
== Global constants
== Part of VstoAddinInstaller
== (https://github.com/bovender/VstoAddinInstaller)
== (c) 2016 Daniel Kraus <bovender@bovender.de>
== Published under the Apache License 2.0
== See http://www.apache.org/licenses
=====================================================================
}

const
  WM_CLOSE = $10;
  MAX_PATH = 250;
  MAX_VERSION = 24; //< highest Office version number to check for.
  MIN_VSTOR_BUILD = 40305; //< minimum required build of VSTO runtime 2010.

{ vim: set ft=pascal sw=2 sts=2 et : }
