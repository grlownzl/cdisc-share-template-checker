import json 
from google.appengine.ext import db

class CheckLog(db.Model):
  opener = db.UserProperty(auto_current_user=True)
  sheet = db.StringProperty(verbose_name="Content Template", required=True)
  date = db.DateTimeProperty(verbose_name="Date Checked", auto_now=True)

class FindingCategory(db.Model):
  name = db.StringProperty(verbose_name="Category Name")
        
class ConsistencyFinding(db.Model):
  checkrun   = db.ReferenceProperty(CheckLog, collection_name="findings")
  template   = db.StringProperty(verbose_name="Content Template Name", required=True)
  tab        = db.StringProperty(verbose_name="Content Sheet Tab", required=True)
  field      = db.StringProperty(verbose_name="Content Sheet Element", required=True)
  column     = db.StringProperty(verbose_name="Column", required=True)
  message    = db.StringProperty(verbose_name="Issue Description", required=True)
  categories = db.ListProperty(db.Key)
  
  def as_dict(self):
    return {'owner' : self.checkrun.opener.email(),
            'date' : self.checkrun.date.isoformat(),
            'template' : self.template,
            'sheet' : self.tab,
            'field' : self.field,
            'column' : self.column,
            'message' : self.message}
  
  def as_list(self):
    if self.field.startswith('-'):
      field = "'%s" % self.field
    else:
      field = self.field
    return [self.checkrun.opener.email(),
            self.checkrun.date.strftime("%Y-%b-%d %H:%M"),
            self.template,
            self.tab,
            field,
            self.column,
            '"%s"' % self.message]

class CodedTerminology(db.Model):
  name = db.StringProperty(verbose_name="Term to be coded", required=True)
  code = db.StringProperty(verbose_name="Assigned C-code")  
  terminology_type = db.StringProperty(verbose_name="Terminology Context")
  #contained_in = db.ListProperty(db.StringProperty)
  
  def is_coded(self):
    return self.code not in [None, "CNEW"]
  
  def as_dict(self):
    return {'name' : self.name, 'code' : self.code, 'terminology_type' : self.terminology_type}