[app]

# Application title
title = Battery Scanner

# Package name
package.name = batteryscanner

# Package domain (reverse DNS)
package.domain = org.batteryscanner

# Source code location
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,ttf,txt,json

# Version
version = 1.0.0

# Requirements (what libraries your app needs)
requirements = python3,kivy,kivymd,openpyxl,pyjnius,android,requests

# Orientation
orientation = portrait

# Android permissions
android.permissions = CAMERA, WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE, INTERNET

# Android API levels
android.api = 33
android.minapi = 21
android.ndk = 25c
android.sdk = 33

# Don't include unneeded files
android.add_src = no

# Log level for debugging
log_level = 2

# Presplash and icon (optional - remove if you don't have these files)
# presplash.filename = %(source.dir)s/presplash.png
# icon.filename = %(source.dir)s/icon.png

[buildozer]
log_level = 2
