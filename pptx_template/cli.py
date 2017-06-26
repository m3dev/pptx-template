# csv2pptx2 - import csv to powerpoint template
# coding=utf-8

import sys
import codecs
import logging
import argparse
import json

from io import open
from pptx import Presentation
from six import iteritems

from pptx_template.core import edit_slide, remove_slide, get_slide, remove_slide_id, remove_all_slides_having_id

def process_slide(ppt, slide, model):
  if model == u"remove":
    remove_slide(ppt, slide)
  else:
    edit_slide(slide, model)

def main():
  if sys.version_info[0] == 2:
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout)

  log = logging.getLogger()
  handler = logging.StreamHandler()
  handler.setLevel(logging.DEBUG)
  log.addHandler(handler)

  parser = argparse.ArgumentParser(description = 'Edit pptx with text replace and csv import')
  parser.add_argument('--template',   help='template pptx file (required)', required=True)
  parser.add_argument('--model',      help='array of model object with JSON format', required=True)
  parser.add_argument('--out',        help='template pptx file (required)', required=True)
  parser.add_argument('--debug',      action='store_true', help='output verbose log')
  opts = parser.parse_args()

  if opts.debug:
    log.setLevel(logging.DEBUG)
  else:
    log.setLevel(logging.WARN)

  ppt = Presentation(opts.template)
  model = json.load(open(opts.model, 'r', encoding='utf-8'))

  slides = model['slides']
  if isinstance(slides, dict):
    for (slide_id, model) in iteritems(slides):
      slide = get_slide(ppt, slide_id)
      remove_slide_id(ppt, slide_id)
      process_slide(ppt, slide, model)
    remove_all_slides_having_id(ppt)
  elif isinstance(slides, list):
    for (model, slide) in zip(slides, ppt.slides):
      process_slide(ppt, slide, model)

  ppt.save(opts.out)


if __name__ == '__main__':
  main()
