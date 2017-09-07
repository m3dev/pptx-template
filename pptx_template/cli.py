# pptx-template - import csv to powerpoint template
# coding=utf-8

import sys
import codecs
import logging
import argparse
import json
import shutil
import os
import tempfile

import openpyxl as xl
from io import open, TextIOWrapper
from pptx import Presentation
from six import iteritems
from itertools import islice

from pptx_template.core import edit_slide, remove_slide, get_slide, remove_slide_id, remove_all_slides_having_id
from pptx_template.xlsx_model import generate_whole_model
from pptx_template import __version__

def process_one_slide(ppt, slide, model, skip_model_not_found = False):
    if model == u"remove":
        remove_slide(ppt, slide)
    else:
        edit_slide(slide, model, skip_model_not_found)


def process_all_slides(slides, ppt, skip_model_not_found = False):
    if isinstance(slides, dict):
        for (slide_id, model) in iteritems(slides):
            slide = get_slide(ppt, slide_id)
            remove_slide_id(ppt, slide_id)
            log.info("Processing slide_id: %s" % slide_id)
            process_one_slide(ppt, slide, model, skip_model_not_found)
        remove_all_slides_having_id(ppt)
    elif isinstance(slides, list):
        for (model, slide) in zip(slides, ppt.slides):
            process_one_slide(ppt, slide, model, skip_model_not_found)


def main():
    parser = argparse.ArgumentParser(description = 'Edit pptx with text replace and csv import')
    parser.add_argument('--template',  help='template pptx file (required)', required=True)
    parser.add_argument('--model',     help='model object file with .json or .xlsx format', required=True)
    parser.add_argument('--out',       help='created pptx file (required)', required=True)
    parser.add_argument('--debug',     action='store_true', help='output verbose log')
    parser.add_argument('--skip-model-not-found', action='store_true', help='skip if specified key is not found in the model')
    opts = parser.parse_args()

    if not len(log.handlers):
        handler = logging.StreamHandler()
        handler.setLevel(logging.DEBUG)
        log.addHandler(handler)

    if opts.debug:
        log.setLevel(logging.DEBUG)
    else:
        log.setLevel(logging.INFO)

    log.info(u"pptx-template version %s" % __version__)

    if opts.model.endswith(u'.xlsx'):
        slides = generate_whole_model(opts.model, {})
    else:
        if opts.model == u'-' and sys.version_info[0] == 3:
            sys.stdin = TextIOWrapper(sys.stdin.buffer, encoding='utf-8')
        with open(opts.model, 'r', encoding='utf-8') if opts.model != u'-' else sys.stdin as m:
            models = json.load(m)
        slides = models[u'slides']

    log.info(u"Loading template pptx: %s" % opts.template)
    ppt = Presentation(opts.template)
    process_all_slides(slides, ppt, skip_model_not_found = opts.skip_model_not_found)

    log.info(u"Writing pptx: %s" % opts.out)
    ppt.save(opts.out)


log = logging.getLogger()

if __name__ == '__main__':
    if sys.version_info[0] == 2:
        reload(sys)
        sys.setdefaultencoding('utf-8')

    main()
