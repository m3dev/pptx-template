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
from io import open
from pptx import Presentation
from six import iteritems
from itertools import islice

from pptx_template.core import edit_slide, remove_slide, get_slide, remove_slide_id, remove_all_slides_having_id
from pptx_template.xlsxMode import generate_whole_model

def process_one_slide(ppt, slide, model):
  if model == u"remove":
    remove_slide(ppt, slide)
  else:
    edit_slide(slide, model)


def process_all_slides(slides, ppt):
  if isinstance(slides, dict):
    for (slide_id, model) in iteritems(slides):
      slide = get_slide(ppt, slide_id)
      remove_slide_id(ppt, slide_id)
      log.info("Processing slide_id: %s" % slide_id)
      process_one_slide(ppt, slide, model)
    remove_all_slides_having_id(ppt)
  elif isinstance(slides, list):
    for (model, slide) in zip(slides, ppt.slides):
      process_one_slide(ppt, slide, model)


def main():
  parser = argparse.ArgumentParser(description = 'Edit pptx with text replace and csv import')
  parser.add_argument('--template',   help='template pptx file (required)', required=True)
  parser.add_argument('--model',      help='model object file with .json or .xlsx format', required=True)
  parser.add_argument('--out',        help='template pptx file (required)', required=True)
  parser.add_argument('--debug',      action='store_true', help='output verbose log')
  opts = parser.parse_args()

  if opts.debug:
    log.setLevel(logging.DEBUG)
  else:
    log.setLevel(logging.INFO)

  log.info(u"Loading template pptx: %s" % opts.template)
  ppt = Presentation(opts.template)

  if opts.model.endswith(u'.xlsx'):
      current_dir = os.getcwd()
      temp_dir = tempfile.mkdtemp()
      try:
          log.info(u"Working in temporary dir:%s ..." % temp_dir)
          xls = xl.load_workbook(opts.model, read_only=True, data_only=True)
          model_sheet = xls['model']
          slides = generate_whole_model(xls, islice(model_sheet.rows, 1, None), {})
          process_all_slides(slides, ppt)
      finally:
          os.chdir(current_dir)
          shutil.rmtree(temp_dir)
  else:
      with open(opts.model, 'r', encoding='utf-8') as f:
          models = json.load(f)
      slides = models[u'slides']
      process_all_slides(slides, ppt)

  log.info(u"Writing pptx: %s" % opts.out)
  ppt.save(opts.out)


log = logging.getLogger()

if __name__ == '__main__':
  if sys.version_info[0] == 2:
    reload(sys)
    sys.setdefaultencoding('utf-8')

  handler = logging.StreamHandler()
  handler.setLevel(logging.DEBUG)
  log.addHandler(handler)
  main()
