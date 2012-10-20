import webapp2

from google.appengine.ext import db
from google.appengine.api import users
from google.appengine.ext.webapp import template

from checker import view, model

class MainPage(webapp2.RequestHandler):
  def get(self):
    user = users.get_current_user()
    if user:
      self.redirect("/content/checker/")
    else:
      self.redirect(users.create_login_url(self.request.uri))

routes = [('/', MainPage),
          ('/content/checker[/]?', view.ContentChecker),
          ('/content/checker/error/(\d+)', view.ContentCheckError),
          ('/content/checker/rules[/]?', view.ContentCheckRules),
          ('/content/checker/findings[/]?', view.ContentCheckReports),
          ('/content/checker/coffee', view.CoffeeTime),
          ('/terminology[/]?', view.BulkTerminologyHandler),
          ('/terminology/upload[/]?', view.BulkTerminologyUploadHandler),
          ('/terminology/codes[/]?', view.BulkTerminologyHandler),
          ('/terminology/code/(\d+)[/]?', view.TerminologyHandler),
          ]

config = {}
config['webapp2_extras.sessions'] = {
    'secret_key': '24d9b21e11612e365cae12282ba11ec661284d3d',
}

app = webapp2.WSGIApplication(routes,
                              config=config,
                              debug=True)