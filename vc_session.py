#! /usr/bin/env python3

# gaptastic -- match open door positions to available responders

import argparse
import logging
import os
import os.path
import re
import datetime
import time
import pprint
import json
import io
import csv
import sys
import random

import requests
import requests_html

from http.cookiejar import LWPCookieJar, Cookie


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

log = logging.getLogger(__name__)


class ScrapeException(Exception):
    pass



def get_session(config, session=None):
    log.debug("in get_session")
    

    cookies = None

    if session == None:
        try:
            cookies = LWPCookieJar(config.COOKIE_FILE)
            #log.debug("before cookie load")
            cookies.load(ignore_discard=True, ignore_expires=True);
            #log.debug("after cookie load")
        except:
            # couldn't read the file; generate new
            log.debug("exception during cookie load")
            cookies = None

    if cookies == None:
        cookies = _refresh_cookies_using_selenium(config)

    if cookies == None:
        log.fatal("Could not log into volunteer connection")
        sys.exit(1)

    if session == None:
        session = requests_html.HTMLSession()

    session.cookies = cookies

    return session



def _refresh_cookies_using_selenium(config):
    log.debug("refreshing authorization cookies via selenium")

    selenium_driver_type = "chrome"

    if selenium_driver_type == "chrome":

        # initialize selenium
        options = webdriver.ChromeOptions()
        options.headless = True

        # this is needed to allow app to run in a container.  One more reason to get rid of the selenium login stuff
        options.add_argument("--no-sandbox")

        # this is needed for headless to work: https://stackoverflow.com/questions/47061662/selenium-tests-fail-against-headless-chrome
        #options.add_argument("--window-size=1280,1024")
        #options.add_argument("--disable-gpu")
        #options.add_argument("--allow-insecure-localhost")

        #options.add_argument("test-type");
        #options.add_argument("enable-strict-powerful-feature-restrictions");
        options.add_argument("disable-geolocation");

        driver = webdriver.Chrome(options=options)

    elif selenium_driver_type == "firefox":

        options = webdriver.FirefoxOptions()
        options.headless = True

        driver = webdriver.Firefox(options=options)

    else:
        log.fatal("Unknown selenium driver type")
        return 1

    cookies = None

    try:
        user = config['SCRAPE_USER']
        password = config['SCRAPE_PASS']

        get_login_selenium(driver, config, user, password)
        cookies = requests_session_with_selenium_cookies(driver, config)

    finally:
        # clean up afterwards
        driver.close()

    return cookies




def get_login_selenium(driver, config, user, password):
    """
    Do the Volunteer Connection login dance in selenium
    """

    driver.get("https://volunteerconnection.redcross.org")

    # wait for redirects to complete
    WebDriverWait(driver, config.WEB_TIMEOUT).until(EC.title_is('Welcome to Volunteer Connection!'))

    WebDriverWait(driver, 5)

    log.debug("before sleep; look for sso-login-form-input-email")
    time.sleep(3);

    element = driver.find_element_by_css_selector('input.sso-login-form-input-email')
    element.clear()
    element.send_keys(user)

    element = driver.find_element_by_css_selector('input.sso-login-form-input-pass')
    element.clear()
    element.send_keys(password)

    element = driver.find_element_by_class_name('sso-login-submit')
    element.click()

    WebDriverWait(driver, config.WEB_TIMEOUT * 2).until(EC.url_contains('?nd=m_home'))

    # make sure there is an admin menu on the page
    # ZZZ: this isn't enough: we need to make sure they have disaster and member privileges
    #time.sleep(1)   # ZZZ: seems to cure race condition on detecting ADMINISTRATION menu item
    #element = driver.find_element_by_link_text('ADMINISTRATION')
    



def requests_session_with_selenium_cookies(driver, config):
    """
    Return a Requests library session object initialized with the cookies from Selenium.

    We have already logged into Volunteer Connection using selenium; use those cookies to
    initialize a Requests session that we will use to download files (Selenium has trouble
    intercepting file downloads)
    """

    cookies = LWPCookieJar(config.COOKIE_FILE)

    selenium_cookies = driver.get_cookies()

    for c in selenium_cookies:
        log.debug(f"selenium cookie: { c }")

        path = c['path']
        path_specified = path != None

        domain = c['domain']
        domain_specified = domain != None
        domain_initial_dot = domain_specified and domain[0] == '.'

        if 'expiry' in c:
            expires = c['expiry'] + 86400 * 365 * 10 # add 10 years to expiry
        else:
            expires = None

        cookie = Cookie(
                version=0,
                name=c['name'],
                value=c['value'],
                port=None,
                port_specified=False,
                discard=False,
                comment=None,
                comment_url=None,
                domain=c['domain'],
                domain_specified=domain_specified,
                domain_initial_dot=domain_initial_dot,
                expires=expires,
                path=path,
                path_specified=path_specified,
                rest={'HttpOnly': c['httpOnly']},
                secure=c['secure'])

        log.debug(f"cookejar cookie: { cookie }\n")

        cookies.set_cookie(cookie)

    cookies.save(ignore_discard=True, ignore_expires=True)
    return cookies

