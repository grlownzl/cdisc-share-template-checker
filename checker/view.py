import webapp2
import os
import json
import csv

from google.appengine.ext.webapp import template
from google.appengine.api import users
from google.appengine.ext import db

from check_content_sheet import ContentSheetChecker
import jinja2
jinja_environment = jinja2.Environment(
    loader=jinja2.FileSystemLoader(os.path.join(os.path.dirname(__file__), "..", "templates")))

import model

class ContentCheckRules(webapp2.RequestHandler):
  """
  Presents the errors
  """
  def get(self):
    checker = ContentSheetChecker()
    template_values = {'rules' : checker.rule_specifications}
    template = jinja_environment.get_template('rules.html')
    self.response.out.write(template.render(template_values))
    
class ContentCheckError(webapp2.RequestHandler):
  """
  Presents the errors
  """
  def get(self, check_id):
    try:
      checklog = model.CheckLog.get_by_id(int(check_id))
    except ValueError:
      # TODO: error
      self.redirect_to("/checker")
    if not checklog:
      # TODO: error
      self.redirect_to("/checker")
    if checklog.findings.count() == 0:
      # TODO: error - no errors exist
      self.redirect_to("/checker")
    else:
      if self.request.get("format"):
        if self.request.get("format") == "csv":
          dest_name = "SCC_%s.csv" % os.path.splitext(checklog.sheet.replace(' ', '_'))[0]
          self.response.headers['Content-type'] = 'text/csv'
          self.response.headers['Content-disposition'] = str("attachment;filename=%s" % dest_name)
          writer = csv.writer(self.response.out)
          writer.writerow("Owner|Date|Template|Sheet|Field|Column|Message".split('|'))
          for issue in [x.as_list() for x in checklog.findings]:
            writer.writerow(issue)
        elif self.request.get("format") == "json":
          self.response.headers["Content-type"] = "application/json"
          self.response.out.write(json.dumps([x.as_dict() for x in checklog.findings]))
      else:
        template_values = {'checklog' : checklog}
        template = jinja_environment.get_template('outcome.html')
        self.response.out.write(template.render(template_values))
    
    
class ContentChecker(webapp2.RequestHandler):

  def get(self):
    path = os.path.join(os.path.dirname(__file__), "..", "templates", 'contentcheck.html')
    template_values = {}
    page = template.render(path, template_values)
    self.response.out.write(page)

  def post(self):
    sheet = self.request.POST["content_sheet"]
    if not hasattr(sheet, "filename"):
      # TODO: Error this
      print "Redirecting: No file" 
      self.redirect("/checker")
    elif sheet.filename.endswith("Template.xlsx"):
      # TODO: Error this
      print "Redirecting: Filename incorrect"
      self.redirect("/checker")
    if users.get_current_user():
      checklog = model.CheckLog(user=users.get_current_user(), sheet=sheet.filename)
      checklog.put()
    else:
      # TODO: Error this
      print "Redirecting: No current user"
      self.redirect("/checker")
    
    # instantiate the checker
    checker = ContentSheetChecker()
    checker.load_from_mem(sheet)
    print checker.exceptions
    if checker.has_issues:
      # log the issues to the datastore
      issues = checker.as_dict()
      for issue in issues:
        # log the issues to the datastore and then query them
        finding = model.ConsistencyFinding(checkrun=checklog,
                                            template=issue.get('template'),
                                            tab=issue.get('sheet'),
                                            field=issue.get('field'),
                                            column=issue.get('column'),
                                            message=issue.get('message'))
        finding.put()
        print "Logged %s" % issue
      self.redirect("/checker/error/%s" % checklog.key().id())
    else:
      self.redirect("/checker/coffee")
      
class ContentGenericRenderer(webapp2.RequestHandler):
  
  def get(self, to_render):
    try:
      template = jinja_environment.get_template('%s.html' % to_render)
    except jinja2.TemplateDoesNotExist:
      self.redirect_to("/checker")
    template_values = {}
    self.response.out.write(template.render(template_values))
