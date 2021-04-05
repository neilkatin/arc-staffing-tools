#! /usr/bin/env python3

import logging

import jinja2


log = logging.getLogger(__name__)



def init():
    env = jinja2.Environment(
            keep_trailing_newline=True,
            auto_reload=False,
            loader=jinja2.FileSystemLoader("."),
            autoescape=jinja2.select_autoescape(['html'])
            )

    return env

def get_template(env, template_file):

    template = env.get_template(template_file)

    return template


