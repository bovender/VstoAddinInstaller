VstoAddinInstaller
===================

[InnoSetup][] script to install and activate Visual Studio Tools for
Office&reg; (VSTO) addins.

Features
--------

- Installs Word and Excel add-ins.
- Checks if Excel or Word is running and can automatically shut it down before
  proceeding with the installation process.
- Can be used with an `/UPDATE` switch to silently shut down and restart Excel
  or Word after the installation.
- Modular structure makes it easy to keep custom configuration separate from
  the core functionality.

The script is based on the installer used by [Daniel's XL Toolbox][].


Obtaining
---------

- If you have [Git for Windows][] installed on your system, you can simply
  clone the repository: <https://github.com/bovender/VstoAddinInstaller.git>
- If you do not have Git, just download the latest [ZIP file][] from
  Github. The file contains a directory "VstoAddinInstaller-master",
  so you can simply unzip it into the downloads folder without
  polluting your files.


Usage
-----

The scipt is divided into several files. The master file,
`vsto-installer.iss`, pulls in several non-configurable files from
the `inc\` subfolder as well as customized configuration files from
the main folder (see below).

To generate an installer, use InnoSetup to compile the master file
`vsto-installer.iss`. Never make changes to this master file; use
custom configuration files instead: To protect you from accidentally
overwriting your personalized configuration with an update from the
Git repository, the distributed configuration files (to be recognized
by the ".dist" contained in the filename) need to be copied from the
`config-dist` folder to the parent folder. Then, rename the files to
remove the ".dist" part from them, and edit these files.

Depending on how much you want to customize the script, you need to
copy and rename just one or several files.


### Most basic scenario ###

The most basic scenario assumes that you have an `.XLAM` and/or an
`.XLA` file, but no other files that need to be installed.

Copy the distributed configuration file `config-dist\config.dist.iss`
to the main folder and rename it to `config.iss`.

Edit the new file `config.iss` and insert the appropriate descriptive
information. By default, the `.XLAM`/`.XLA` files are expected in a
`source\` folder, but this can be adjusted in the `config.iss` file
too.

__Important:__ When you first edit this file, you *must* create a
global unique ID (GUID) for your addin. You will easily identify the
line in the default configuration file where this information is
needed.  InnoSetup has a "Generate GUID" command in the "Tools" menu.

When you are done editing, save the file, then right-click on the
`vsto-installer.iss` file and choose "Compile" from the context menu
(if you do not see a "Compile" command in the context menu, check that
you have actually installed [InnoSetup]).

Alternatively, double-click on `vsto-installer.iss`, which will start
InnoSetup with the file loaded.

The installer will be written to the `deploy\` folder by default. This
can be changed in the `config.iss` file.


### Advanced configuration ###

If you need more advanced configuration, copy and rename one or
several of the following configuration files from the `config-dist`
folder to the main folder:
- `lanuages.dist.iss` and `messages.dist.iss` to add more languages.
- `tasks.dist.iss` to define custom tasks.

Always remember to remove the `.dist` from the file names after
copying them to the main folder.



Further information
-------------------

For background information, see
<http://xltoolbox.net/blog/2015-01-30-net-vsto-add-ins-getting-prerequisites-right.html>.


Notice
------

This script is based on the related [ExcelAddinInstaller][].


License
-------

Published under the [Apache License, Version 2.0](LICENSE).

        Copyright (C) 2015 Daniel Kraus <http://github.com/bovender>

        Licensed under the Apache License, Version 2.0 (the "License");
        you may not use this file except in compliance with the License.
        You may obtain a copy of the License at

        http://www.apache.org/licenses/LICENSE-2.0

        Unless required by applicable law or agreed to in writing, software
        distributed under the License is distributed on an "AS IS" BASIS,
        WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
        See the License for the specific language governing permissions and
        limitations under the License.

Microsoft®, Windows®, Office, and Excel® are either registered
trademarks or trademarks of Microsoft Corporation in the United States
and/or other countries.


[InnoSetup]: http://www.jrsoftware.org/isinfo.php
[Daniel's XL Toolbox]: http://xltoolbox.net
[ZIP file]: https://github.com/bovender/VstoAddinInstaller/archive/master.zip
[Git for Windows]: http://git-scm.com/downloads

<!-- vim: set tw=70 ts=4 :-->
