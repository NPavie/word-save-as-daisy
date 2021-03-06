---
title: Download SaveAsDAISY Word Addin
layout: my-default
---
SaveAsDAISY is available for Windows only. We provide an installer for Microsoft Office 64-bit edition and an installer that detects the installed version of Office. Choose the latter if you don't know whether your Office is the 32 or 64-bit edition.


## Latest version: 2.6.1 beta - Minor update 2 (released on January 27, 2020)

- [Download universal installer](https://github.com/daisy/word-save-as-daisy/releases/download/v2.6.1.2-beta/SaveAsDAISYInstaller.exe)
- [Download 64-bit installer](https://github.com/daisy/word-save-as-daisy/releases/download/v2.6.1.2-beta/SaveAsDAISYInstaller_Office64bits.exe)

## Report issues

If you have an issue with the installers or with the add-in, please contact the development team by mail to [daisy-pipeline@mail.daisy.org](mailto:daisy-pipeline@mail.daisy.org).

Any construcive feedback is also welcome to help us improve the add-in :)

# Changelog

# 2.6.1 beta - Minor update 2 (January 28, 2021)

Minor update including
- bugfix for [#2](https://github.com/daisy/word-save-as-daisy/issues/2) in test phase
- bugfix for [#3](https://github.com/daisy/word-save-as-daisy/issues/3) in test phase
- Documentation updated (new manual and authoring guidelines, including integrated help file)
- Added a document file name check pass before launching conversion (to let the user manually rename his file if wanted)
- Fixed an installer issue where previous version of packages could be used instead of the new ones
- Name and logos updated in software and installers


# 2.6.1 beta - Minor update 1 (January 14, 2021)

Minor update including
- bugfix for [#1](https://github.com/daisy/word-save-as-daisy/issues/1) in test phase
- Removed the Accessibility Checker shortcut in the addin ribbon for office 2010 and later
Users are advised to validated using the accessibility checker before launching a conversion in the conversion setting form, with a link to the checker documentation.
- Updated the AdoptOpenJDK java 8 runtime provided with the embedded DAISY pipeline lite

## Feedbacks

- File not found errors with partial file names reported missing:
Some characters like spaces and commas may lead to "File not found" errors in the conversion process. Replace those characters by underscore "_" and launch the conversion again.

# 2.6.1 beta (December 18, 2020)

## Installer update

- Office XP is removed from the list of supported versions of Office,
- The list of officially supported versions of office now goes from Office 2003 to Office 2016/365,
  - For newer version, a warning is raised to let the user decide if the installation must continue or not.
- For Windows x64, a unified installer is provided for Office 32bits and 64Bits
- For undetected Office installations (OEM or Windows store install of Office are preventing version detection):
  - For office 32bits, the unified installer can be used,
  - For office 64bits, a separate installer is provided.
- The default installation directory is now in program files/DAISY/Save-as-DAISY Word Addin,
- The latest version available of the DAISY pipeline 1 is included in the addin installer, and is shipped with an integrated java runtime for Windows x64

## Features

- The Word validation button is being replaced by a shortcut to the Microsoft Accessibility Checker, available in office 2010 and newer.
  - The previous word validation process is only kept for word 2007 and 2003 where the accessibility checker is not available,
- The conversion of word documents to DAISY XML now includes a post-process pass:
  - User can decide to apply cleanups and sentences detection in the conversion process options

## Feedbacks

- An issue has been reported with document containing shapes or images, where the conversion would not launch. The cause has been identified and is being worked on.
- A user reported a "Not able to find directory to install" warning when installing the addin for word 32 bits on x64 system. The install directory was found after going one step back and clicking again on next.
