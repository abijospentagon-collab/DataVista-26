# -*- coding: utf-8 -*-
"""
DATA VISTA '26 — Email Configuration
=====================================
Fill in your Gmail credentials below to enable confirmation emails.

HOW TO GET A GMAIL APP PASSWORD:
  1. Go to: https://myaccount.google.com/security
  2. Enable "2-Step Verification" (if not already on)
  3. Go to: https://myaccount.google.com/apppasswords
  4. Create a new App Password -> select "Mail" -> "Windows Computer"
  5. Copy the 16-character password (spaces don't matter) and paste below

Set MAIL_ENABLED = False to disable emails (registrations still work).
"""

import os

MAIL_ENABLED  = True
MAIL_SERVER   = 'smtp.gmail.com'
MAIL_PORT     = 587

MAIL_USERNAME = (os.environ.get('MAIL_USERNAME') or 'datavista2026@gmail.com').strip()
MAIL_PASSWORD = (os.environ.get('MAIL_PASSWORD') or 'rlie zhta ifed uvxn').strip()

MAIL_FROM     = f"DATA VISTA '26 <{MAIL_USERNAME}>"
