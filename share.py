import webapp2

from google.appengine.ext import db
from google.appengine.api import users
from google.appengine.ext.webapp import template

from checker import view, model

class MainPage(webapp2.RequestHandler):
  def get(self):
    user = users.get_current_user()
    if user:
      self.redirect("/checker/")
    else:
      self.redirect(users.create_login_url(self.request.uri))

routes = [('/', MainPage),
          ('/checker[/]?', view.ContentChecker),
          ('/checker/error/(\d+)', view.ContentCheckError),
          ('/checker/rules[/]?', view.ContentCheckRules),
          ('/checker/findings[/]?', view.ContentCheckReports),
          ('/checker/(.+)', view.ContentGenericRenderer),
          ]

app = webapp2.WSGIApplication(routes,
                              debug=True)