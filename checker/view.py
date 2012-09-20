import webapp2
from webapp2_extras import sessions
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

class BaseHandler(webapp2.RequestHandler):

  def dispatch(self):
    # Get a session store for this request.
    self.session_store = sessions.get_store(request=self.request)
    try:
      # Dispatch the request.
      webapp2.RequestHandler.dispatch(self)
    finally:
      # Save all sessions.
      self.session_store.save_sessions(self.response)

  @webapp2.cached_property
  def session(self):
    # Returns a session using the default cookie key.
    return self.session_store.get_session()    
    
  MESSAGE_KEY = '_flash_message'
  def add_message(self, message, level=None):
		self.session.add_flash(message, level, BaseHandler.MESSAGE_KEY)
  
  def get_messages(self):
  	return self.session.get_flashes(BaseHandler.MESSAGE_KEY)
  
  def render_jinja(self, name, data):
    data['logout_url']=users.create_logout_url(self.request.uri)
    data['errors'] = self.get_messages()
    try:
      template = jinja_environment.get_template('%s.html' % name)
    except jinja2.TemplateNotFound:
      return self.redirect("/checker")
    self.response.out.write(template.render(data))
        
class ContentCheckRules(BaseHandler):
  """
  Presents the errors
  """
  def get(self):
    checker = ContentSheetChecker()
    template_values = {'rules' : checker.rule_specifications}
    self.render_jinja("rules", template_values)

class ContentCheckReports(BaseHandler):
  
  def get(self):
    checklogs = model.CheckLog.all()
    self.render_jinja("findings", {"checklogs" : checklogs})
        
class ContentCheckError(BaseHandler):
  """
  Presents the errors
  """
  def get(self, check_id):
    try:
      checklog = model.CheckLog.get_by_id(int(check_id))
    except ValueError:
      # TODO: error
      self.add_message("Can't convert supplied id %s" % check_id)
      self.render_jinja("contentchecker", {})
    if not checklog:
      # TODO: error
      self.add_message("No matching CheckLog with id %s" % check_id)
      self.render_jinja("contentchecker", {})
    if checklog.findings.count() == 0:
      # TODO: error - no errors exist
      self.add_message("CheckLog with id %s has no issues" % check_id)
      self.render_jinja("contentchecker", {})
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
          self.add_message("Format %s not recognised" % self.request.get('format'))
          return self.redirect("/checker/errors/%s" % checker.key().id())
      else:
        template_values = {'checklog' : checklog}
        self.render_jinja("outcome", template_values)    
    
class ContentChecker(BaseHandler):

  def get(self):
    template_values = {}
    self.render_jinja("contentcheck", template_values)    


  def post(self):
    sheet = self.request.POST["content_sheet"]
    if not hasattr(sheet, "filename"):
      # TODO: Error this
      self.add_message("No file")
      self.render_jinja("contentcheck", {})
    elif sheet.filename.endswith("Template.xlsx"):
      # TODO: Error this
      self.add_message("File name doesn't match pattern")
      self.render_jinja("contentcheck", {})
    if users.get_current_user():
      checklog = model.CheckLog(user=users.get_current_user(), sheet=sheet.filename)
      checklog.put()
    else:
      # TODO: Error this
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
      self.redirect("/checker/error/%s" % checklog.key().id())
    else:
      self.redirect("/checker/coffee")
      
class ContentGenericRenderer(BaseHandler):
  
  def get(self, to_render):
    template_values = {}
    self.render_jinja(to_render, template_values)
