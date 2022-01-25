#-*- coding:utf-8 -*-
# Part of ChamsERP. See LICENSE file for full copyright and licensing details.
import time, re, logging, gc, smtplib, csv, base64, requests, json, os, zipfile
from itertools import compress
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from decimal import *
from sets import Set
from cStringIO import StringIO
from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx
import xlsxwriter
import xlsx_report_live, shutil

from openerp import api, models, netsvc, registry, _
from openerp.osv import fields, osv, orm
from openerp.tools import DEFAULT_SERVER_DATE_FORMAT
from openerp.tools import DEFAULT_SERVER_DATETIME_FORMAT
from openerp.exceptions import ValidationError
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import Encoders
from email.mime.text import MIMEText
from xlrd import open_workbook
from docxtpl import DocxTemplate
from xlsx_report_live import *
from io import BytesIO

_logger = logging.getLogger(__name__)
#_logger2 = logging.getLogger(__name__)

REPORTS_DIR = '/odoo/odoo9/reports/'

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
 
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
 
    return False
   
def last_day_of_month(y, m):
    '''
    Returns an integer representing the last day of the month, given
    a year and a month.
    '''
 
    # Algorithm: Take the first day of the next month, then count back
    # ward one day, that will be the last day of a given month. The 
    # advantage of this algorithm is we don't have to determine the 
    # leap year.
 
    m += 1
    if m == 13:
        m = 1
        y += 1
 
    first_of_next_month = date(y, m, 1)
    last_of_this_month = first_of_next_month + timedelta(-1)
    return last_of_this_month.day

class res_users(osv.osv):
    _inherit = 'res.users'
    
    _columns = {
        'domain_mdas': fields.many2many('hr.department', 'rel_user_domain_mdas', 'user_id', 'department_id', 'Domain MDAs',),
        'domain_tcos': fields.many2many('ng.state.payroll.tco', 'rel_user_domain_tcos', 'user_id', 'tco_id', 'Domain TCOs',),
        'domain_tco_types': fields.one2many('ng.state.payroll.tcodomain','user_id','Domain TCOs'),
    }

class res_bank(osv.osv):
    '''
    Bank
    '''
    _inherit = "res.bank"
    _description = "Bank"

    _columns = {
        'code': fields.char('Bank Code', help='Bank Code', required=True),
    }
        
class hr_holidays(osv.osv):
    '''
    HR Leave
    '''
    _inherit = "hr.holidays"
    _description = "HR Leave"
    
    @api.model
    def add_leave(self, vals):
        _logger.info("add_leave - user_id=%d, vals=%s", self.env.uid, vals)
        #Get employee with user_id
        employee_obj = self.env['hr.employee'].search([('user_id', '=', self.env.uid)])
        if employee_obj:
            employee_id = employee_obj[0].id
            holiday_type = 'employee'
            add_type = 'remove'
            state = 'draft'
            name = vals.get('description')
            holiday_status_id = vals.get('leave_type')
            no_days = vals.get('days')
            insert_phrase = "(" + str(employee_id) + ",'" + name + "'," + str(holiday_status_id) + ",'" + holiday_type + "','" + state + "','" + add_type + "',-" + str(no_days) + ")"
            _logger.info("add_leave - insert_phrase=%s", insert_phrase)
            self.env.cr.execute("insert into hr_holidays (employee_id,name,holiday_status_id,holiday_type,state,type,number_of_days) values " + insert_phrase)
            return self.env.uid
        else:
            return "Wrong login credentials; please login"
                           
    @api.model
    def list_statuses(self, context=None):
        _logger.info("list_leave_statuses")
        self.env.cr.execute("select id,name from hr_holidays_status")
        item_lines = self.env.cr.fetchall()
        return item_lines
    
    @api.model
    def list_leave_items(self, context=None):
        _logger.info("list_leave_items - %d", self.env.uid)
        employee_obj = self.env['hr.employee'].search([('user_id', '=', self.env.uid)])
        item_lines = []
        if employee_obj:
            employee_id = employee_obj[0].id
            self.env.cr.execute("select id,name,(number_of_days || 'days'),upper(state) from hr_holidays where employee_id=" + str(employee_id))
            item_lines = self.env.cr.fetchall()
        return item_lines    
    
class hr_employee(osv.osv):
    '''
    Employee
    '''

    _inherit = "hr.employee"
    _description = 'Employee'

    _columns = {
        'sinid': fields.char('Pension PIN', help='Pension PIN'),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'resolved_earn_dedt': fields.boolean('Resolve Earnings/Deductions', help='Resolve Earnings/Deductions', required=False),
        'employee_no': fields.char('Employee Number', help='Employee Number'),
        'bvn': fields.char('BVN', help='Bank Verification Number'),
        'school_emp_id': fields.char('School Employee ID', help='School Employee ID', required=False),
        'bank_account_no': fields.char('Bank Account', help='Bank Account Number'),
        'job_description': fields.text('Job Description', help='Job Description', required=False),
	'birthday': fields.date('Birth Date', help='Date of Birth'),
        'hire_date': fields.date('Hire Date', help='Date of Hire'),
        'confirmation_date': fields.date('Confirmation Date', help='Date of Confirmation'),
        'retirement_due_date': fields.date('Retirement-Due Date', help='Retirement-Due Date'),
        'last_promotion_date': fields.date('Last Promotion Date', help='Last Promotion Date'),
        'next_promotion_date': fields.date('Next Promotion Date', help='Next Promotion Date'),
        'lga_id': fields.many2one('ng.state.payroll.lga', 'LGA'),
        'pfa_id': fields.many2one('ng.state.payroll.pfa', 'PFA'),
        'mfb_id': fields.many2one('ng.state.payroll.mfb', 'MFB'),
        'mfb_account': fields.char('MFB Account No', help='MFB Account Number'),
        'school_id': fields.many2one('ng.state.payroll.school', 'School', required=False),
        'paycategory_id': fields.many2one('ng.state.payroll.paycategory', 'Pay Category'),
        'payscheme_id': fields.many2one('ng.state.payroll.payscheme', 'Pay Scheme'),
        'level_id': fields.many2one('ng.state.payroll.level', 'Grade'),
        'level_id_leave_allowance': fields.many2one('ng.state.payroll.level', 'January Grade for Leave Bonus Computation'),
        'retirement_index': fields.selection([
            ('dofb', 'Date of Birth'),
            ('dofa', 'Date of First Appointment'),
        ], 'Retirement Index'),
        'grade_level': fields.selection([
            (1, 'GL-1'),
            (2, 'GL-2'),
            (3, 'GL-3'),
            (4, 'GL-4'),
            (5, 'GL-5'),
            (6, 'GL-6'),
            (7, 'GL-7'),
            (8, 'GL-8'),
            (9, 'GL-9'),
            (10, 'GL-10'),
            (12, 'GL-12'),
            (13, 'GL-13'),
            (14, 'GL-14'),
            (15, 'GL-15'),
            (16, 'GL-16'),
            (17, 'GL-17'),
            (18, 'GL-18'),
            (19, 'GL-19'),
            (20, 'GL-20'),
        ], 'Grade Level'),
        'grade_step': fields.selection([
            (1, 'Step-1'),
            (2, 'Step-2'),
            (3, 'Step-3'),
            (4, 'Step-4'),
            (5, 'Step-5'),
            (6, 'Step-6'),
            (7, 'Step-7'),
            (8, 'Step-8'),
            (9, 'Step-9'),
            (10, 'Step-10'),
            (11, 'Step-11'),
            (12, 'Step-12'),
            (13, 'Step-13'),
            (14, 'Step-14'),
            (15, 'Step-15'),
            (16, 'Step-16'),
            (17, 'Step-17'),
            (18, 'Step-18'),
            (19, 'Step-19'),
            (20, 'Step-20'),
        ], 'Grade Step'),
        'title_id': fields.many2one('res.partner.title', 'Title'),
        'status_id': fields.many2one('ng.state.payroll.status', 'Employee Status'),
        'bank_id': fields.many2one('res.bank', string='Bank'),
        'contract_id': fields.many2one('hr.contract', 'Contract', required=False),
        'designation_id': fields.many2one('ng.state.payroll.designation', 'Designation', required=False),
        'disciplinary_actions': fields.one2many('ng.state.payroll.disciplinary', 'employee_id', 'Disciplinary Actions'),
        'promotions': fields.one2many('ng.state.payroll.promotion', 'employee_id', 'Promotions'),
        'salary_items': fields.one2many('ng.state.payroll.payroll.item', 'employee_id', 'Salary History', compute='_compute_salary_items'),
        'pension_items': fields.one2many('ng.state.payroll.pension.item', 'employee_id', 'Pension History', compute='_compute_pension_items'),
        'payment_items': fields.one2many('ng.state.payroll.scenario.payment', 'employee_id', 'Payment History'),
        'payment2_items': fields.one2many('ng.state.payroll.scenario2.payment', 'employee_id', 'Payment History'),
        'query_items': fields.one2many('ng.state.payroll.query', 'employee_id', 'Query History'),
        'pensiontype_id': fields.many2one('ng.state.payroll.pensiontype', 'Pension Type', required=False),
        'tco_id': fields.many2one('ng.state.payroll.tco', 'TCO', required=False),
        'pensionfile_no': fields.char('Pension File', help='Pension File Number'),
        'annual_pension': fields.float('Annual Pension', help='Annual Pension'),
        'loan_items': fields.many2many('ng.state.payroll.loan', 'rel_employee_loan', 'employee_id', 'loan_id', 'Loans'),
        'standard_earnings': fields.many2many('ng.state.payroll.earning.standard', 'rel_employee_std_earning', 'employee_id', 'earning_id', 'Standard Earnings'), 
        'standard_deductions': fields.many2many('ng.state.payroll.deduction.standard', 'rel_employee_std_deduction', 'employee_id', 'deduction_id', 'Standard Deductions'), 
        'certifications': fields.many2many('ng.state.payroll.certification', 'rel_employee_certification', 'employee_id', 'certification_id', 'Certifications'), 
        'trainings': fields.many2many('ng.state.payroll.training.history', 'rel_employee_training', 'employee_id', 'training_id', 'Certifications'), 
        'nonstd_earnings': fields.one2many('ng.state.payroll.earning.nonstd', 'employee_id', 'Nonstandard Earnings'),
        'nonstd_deductions': fields.one2many('ng.state.payroll.deduction.nonstd', 'employee_id', 'Nonstandard Deduction'),
        'employee_earnings': fields.one2many('ng.state.payroll.earning.employee', 'employee_id', 'Employee Earnings'),
        'employee_deductions': fields.one2many('ng.state.payroll.deduction.employee', 'employee_id', 'Employee Deduction'),
        'pension_arrears': fields.one2many('ng.state.payroll.arrears.pension', 'employee_id', 'Pensioner Arrears'),
        'hr_events': fields.one2many('ng.state.payroll.hrevent', 'employee_id', 'HR Events'),
    }
    
    _sql_constraints = [
        ('bvn_unique', 'unique(bvn)', 'An employee with this BVN already exists!')
    ]
    
    @api.onchange('department_id')
    def school_id_update(self):
        return {'domain': {'school_id': [('org_id','=',self.department_id.id)] }}
    
    @api.onchange('designation_id')
    def cadre_update(self):
        return {'domain': {'designation_id': [('cadre_id','=',self.designation_id.cadre_id.id)] }}
        
    @api.multi
    def _compute_salary_items(self):
        self.salary_items = self.env['ng.state.payroll.payroll.item'].search([('active','=',True),('employee_id.id','=',self.id),('payroll_id.state','=','closed')])
        
    @api.multi
    def _compute_pension_items(self):
        self.salary_items = self.env['ng.state.payroll.pension.item'].search([('active','=',True),('employee_id.id','=',self.id),('payroll_id.state','=','closed')])
        
class hr_department(osv.osv):
    _name = "hr.department"
    _description = "Organization"
    _inherit = 'hr.department'

    _columns = {
        'name': fields.char('MDA', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'company_id': fields.many2one('res.company', 'Organization', select=True, required=False),
        'parent_id': fields.many2one('hr.department', 'Parent MDA', select=True),
        'orgtype_id': fields.many2one('ng.state.payroll.orgtype', 'MDA Type', select=True),
        'child_ids': fields.one2many('hr.department', 'parent_id', 'Child MDAs'),
        'member_ids': fields.one2many('hr.employee', 'department_id', 'Employees', readonly=True),
        'school_ids': fields.one2many('ng.state.payroll.school', 'org_id', 'Schools'),
    } 
    
class ng_state_payroll_school(models.Model):
    _name = "ng.state.payroll.school"
    _description = "School"

    _columns = {
        'name': fields.char('School Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'org_id': fields.many2one('hr.department', 'Parent Organization', select=True),
        'teacher_ids': fields.one2many('hr.employee', 'school_id', 'Teachers'),
    }  
    
class ng_state_payroll_cadre(models.Model):
    _name = "ng.state.payroll.cadre"
    _description = "Cadre"

    _columns = {
        'name': fields.char('Cadre Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'designation_ids': fields.one2many('ng.state.payroll.designation', 'cadre_id', 'Designations', readonly=True),
    }  

    _defaults = {
        'active': True,
    }             
    
class ng_state_payroll_designation(models.Model):
    _name = "ng.state.payroll.designation"
    _description = "Designation"

    _columns = {
        'name': fields.char('Designation Name', required=True),
        'code': fields.char('Designation Code', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'cadre_id': fields.many2one('ng.state.payroll.cadre', 'Cadre', required=True),
        'paygrade_id': fields.many2one('ng.state.payroll.paygrade', 'Grade Level', required=False),
    }  

    _defaults = {
        'active': True,
    }
                     
    _sql_constraints = [
        ('code', 'unique(code)', 'Code already exists; must be unique!')
    ]
    
    @api.model
    def create(self, vals):
        vals['code'] = 'D' + str(self.env['ir.sequence'].next_by_code('ng.state.payroll.designation')).zfill(6)
        res = super(ng_state_payroll_designation, self).create(vals)
            
        return res  
        
    @api.multi
    @api.constrains('code')
    def validate_code(self):
        designation_singleton = self.env['ng.state.payroll.designation']
        
        for obj in self:
            designation_ids = designation_singleton.search([('code' == obj.code)])
            if designation_ids:
                raise ValidationError("Code already exists; must be unique: %s" % obj.code)
        
        return True    
    
    @api.multi
    def name_get(self):
 
        data = []
        for d in self:
            display_value = ''
            display_value += d.name
            display_value += ' - '
            display_value += d.cadre_id.name
            data.append((d.id, display_value))
            
        return data
               
class ng_state_payroll_relief(models.Model):
    '''
    Relief
    '''
    _name = "ng.state.payroll.relief"
    _description = 'Relief'

    _columns = {
        'name': fields.char('Relief', help='Relief Name', required=True),
        'code': fields.char('Code', help='Relief Code', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
    }  

    _defaults = {
        'active': True,
    }             
           
class ng_state_payroll_pensiontype(models.Model):
    '''
    Pension Type
    '''
    _name = "ng.state.payroll.pensiontype"
    _description = 'Pension Type'

    _columns = {
        'name': fields.char('Type', help='Pension Type Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
    }  

    _defaults = {
        'active': True,
    }             

class ng_state_payroll_orgtype(models.Model):
    '''
    Organization Type
    '''
    _name = "ng.state.payroll.orgtype"
    _description = 'Organization Type'

    _columns = {
        'name': fields.char('Type', help='Organization Type Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
    }       

    _defaults = {
        'active': True,
    }             

class ng_state_payroll_tcodomain(models.Model):
    '''
    TCO Domain
    '''
    _name = "ng.state.payroll.tcodomain"
    _description = 'Domain'

    _columns = {
        'user_id': fields.many2one('res.users', 'Pension Administrator', required=True, domain="[('groups_id.name','=','Payroll Administrator')]"),
        'tco_id': fields.many2one('ng.state.payroll.tco', 'TCO', required=True),
        'pensiontype_id': fields.many2one('ng.state.payroll.pensiontype', 'Pension Type', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
    }

    _defaults = {
        'active': True,
    }    
    
class ng_state_payroll_tco(models.Model):
    '''
    Treasury Cash Office
    '''
    _name = "ng.state.payroll.tco"
    _description = 'Treasury Cash Office'

    _columns = {
        'name': fields.char('TCO Name', help='Treasury Cash Office', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
    }

    _defaults = {
        'active': True,
    }             

class ng_state_payroll_lga(models.Model):
    '''
    Local Government Area
    '''
    _name = "ng.state.payroll.lga"
    _description = 'Local Government Areas'

    _columns = {
        'name': fields.char('LGA Name', help='Local Government Area', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'country_state': fields.many2one('res.country.state', 'Country State', required=True),
    }

    _defaults = {
        'active': True,
    }             

class ng_state_payroll_pfa(models.Model):
    '''
    Pension Fund Administrator
    '''
    _name = "ng.state.payroll.pfa"
    _description = 'Pension Fund Administrator'

    _columns = {
        'name': fields.char('PFA Name', help='Pension Fund Administrator Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
    }

    _defaults = {
        'active': True,
    }             
    
class ng_state_payroll_mfb(models.Model):
    '''
    Micro-Finance Bank
    '''
    _name = "ng.state.payroll.mfb"
    _description = 'Micro-Finance Bank'

    _columns = {
        'name': fields.char('MFB Name', help='Micro-Finance Bank', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'account_no': fields.char('Account Number', help='Account Number', required=True),
        'parent_bank_id': fields.many2one('res.bank', string='Bank', required=True),
        'email': fields.char('MFB Email', help='MFB Email Address', required=True),
    }

    _defaults = {
        'active': True,
    }             
          
class ng_state_payroll_status(models.Model):
    '''
    Employee Status
    '''
    _name = "ng.state.payroll.status"
    _description = 'Employee Status'

    _columns = {
        'name': fields.char('Name', help='Employee Status', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
    }    

    _defaults = {
        'active': True,
    }             
    
class ng_state_payroll_paycategory(models.Model):
    '''
    Pay Category
    '''
    _name = "ng.state.payroll.paycategory"
    _description = 'Pay Category'

    _columns = {
        'name': fields.char('Name', help='Category Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'department': fields.many2one('hr.department', 'MDA', required=True),
    }

    _defaults = {
        'active': True,
    }             

class ng_state_payroll_level(models.Model):
    '''
    Pay Grade/Level
    '''
    _name = "ng.state.payroll.level"
    _description = 'Pay Grade/Level'
    
    _columns = {
        'step': fields.integer('Step Name', help='Step Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'paygrade_id': fields.many2one('ng.state.payroll.paygrade', 'Pay Grade', required=True),
    }

    _defaults = {
        'active': True,
    }
                 
    _rec_name = 'step'
    
    @api.multi
    def name_get(self):
 
        data = []
        for d in self:
            display_value = ''
            display_value += str(d.paygrade_id.level).zfill(2)
            display_value += '.'
            display_value += str(d.step).zfill(2)
            data.append((d.id, display_value))
            
        return data
        
class ng_state_payroll_payscheme(models.Model):
    '''
    Pay Scheme
    '''
    _name = "ng.state.payroll.payscheme"
    _description = 'Pay Scheme'        

    _columns = {
        'name': fields.char('Name', help='Pay Scheme', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'use_dob': fields.boolean('Use DofB', help='Use birth date for retirement date computation', required=True),
        'use_dofa': fields.boolean('Use DofA', help='Use appointment date for retirement date computation', required=True),
        'retirement_age': fields.integer('Retirement Age', help='Expected retirement age', required=True),
        'service_years': fields.integer('Service Years', help='Number of years at which retirement is due', required=True),
        'employee_category': fields.selection([
            ('public', 'Public Servant'),
            ('political', 'Political Officer'),
            ('contract', 'Contract Employee'),
            ], 'Employee Category', required=False),
        
    }

    _defaults = {
        'active': True,
        'use_dob': True,
        'use_dofa': True,
        'employee_category': 'public',
    }             
    
class ng_state_payroll_paygrade(models.Model):
    '''
    Pay Grade
    '''
    _name = "ng.state.payroll.paygrade"
    _description = 'Pay Grade'

    _columns = {
        'active': fields.boolean('Active', help='Active Status', required=True),
        'level': fields.integer('Grade Level', help='Grade Level', required=True),
        'payscheme_id': fields.many2one('ng.state.payroll.payscheme', 'Pay Scheme', required=True),
        'gross_ceiling': fields.float('Gross Ceiling', help='Gross Ceiling', required=False),
    } 

    _defaults = {
        'active': True,
        'gross_ceiling': 1000000000.0,
    }             

    def name_get(self, cr, uid, ids, context=None):
        if not ids:
            return []
        res = []
        for r in self.read(cr, uid, ids, ['id', 'level', 'payscheme_id'], context):
            aux = ('(')
            if r['level']:
                aux += ('Grade Level - ' + str(r['level']) + ', ') # same question
    
            if r['payscheme_id']:
                aux += ('Pay Scheme - ' + r['payscheme_id'][1]) # same question
            aux += (')')
    
            # aux is the name items for the r['id']
            res.append((r['id'], aux))  # append add the tuple (r['id'], aux) in the list res
    
        return res

        #Open create form with current month date range
    def name_search(self, cr, user, name='', args=None, operator='ilike', context=None, limit=100):
        if not args:
            args = []
        if name:
            ids = self.search(cr, user, [('payscheme_id.name','=ilike',name)]+ args, limit=limit, context=context)
            if not ids:
                # Do not merge the 2 next lines into one single search, SQL search performance would be abysmal
                # on a database with thousands of matching products, due to the huge merge+unique needed for the
                # OR operator (and given the fact that the 'name' lookup results come from the ir.translation table
                # Performing a quick memory merge of ids in Python will give much better performance
                ids = set()
                ids.update(self.search(cr, user, args + [('level',operator,name)], limit=limit, context=context))
                    #End
                ids = list(ids)
            if not ids:
                ptrn = re.compile('(\[(.*?)\])')
                res = ptrn.search(name)
                if res:
                    ids = self.search(cr, user, [('id','=', res.group(2))] + args, limit=limit, context=context)
        else:
            ids = self.search(cr, user, args, limit=limit, context=context)
        result = self.name_get(cr, user, ids, context=context)
        return result
    
class ng_state_payroll_arrears_pension(models.Model):
    '''
    Pension Arrears
    '''
    _name = "ng.state.payroll.arrears.pension"
    _description = 'Pension Arrears'

    _columns = {
        'name': fields.char('Name', help='Arrears Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'amount': fields.float('Amount', help='Amount', required=True),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'calendar_id': fields.many2one('ng.state.payroll.calendar', 'Calendar', required=True),
    }

    _defaults = {
        'active': True,
    }
    
    @api.onchange('name')
    def name_update(self):
        if self.name:
            self.name = self.name.upper()
	    self.name = self.name.strip()
    
class ng_state_payroll_deduction_pension(models.Model):
    '''
    Pension Deduction
    '''
    _name = "ng.state.payroll.deduction.pension"
    _description = 'Pension Deduction'

    _columns = {
        'name': fields.char('Name', help='Deduction Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'fixed': fields.boolean('Fixed Amount', help='Fixed Amount'),
        'amount': fields.float('Amount', help='Amount', required=True),
        'whitelist_ids': fields.many2many('hr.employee', 'rel_deduction_pension_whitelist', 'deduction_id', 'employee_id', 'Whitelist', domain="[('status_id.name','=','PENSIONED'),('active','=',True)]",),
        'blacklist_ids': fields.many2many('hr.employee', 'rel_deduction_pension_blacklist', 'deduction_id', 'employee_id', 'Blacklist', domain="[('status_id.name','=','PENSIONED'),('active','=',True)]"),
    }

    _defaults = {
        'active': True,
    }             
    
    @api.onchange('name')
    def name_update(self):
        if self.name:
            self.name = self.name.upper()
	    self.name = self.name.strip()
        
class ng_state_payroll_deduction_standard(models.Model):
    '''
    Standard Deduction
    '''
    _name = "ng.state.payroll.deduction.standard"
    _description = 'Standard Deduction'

    _columns = {
        'name': fields.char('Name', help='Deduction Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'fixed': fields.boolean('Fixed Amount', help='Fixed Amount'),
        'relief': fields.boolean('Relief', help='Forming part of CRA Relief'),
        'income_deduction': fields.boolean('Income Deduction', help='Deducted from Income before CRA 20% calculation'),
        'amount': fields.float('Amount', help='Amount', required=True),
        'payscheme_id': fields.many2one('ng.state.payroll.payscheme', 'Pay Scheme', required=True),
        'level_id': fields.many2one('ng.state.payroll.level', 'Grade', required=True),
        'derived_from': fields.many2one('ng.state.payroll.earning.standard', 'Derived From', required=False),
    }

    _defaults = {
        'active': True,
    }             
    
    @api.onchange('name')
    def name_update(self):
        if self.name:
            self.name = self.name.upper()
	    self.name = self.name.strip()
        
    @api.onchange('level_id','payscheme_id')
    def derived_from_id_update(self):
        return {'domain': {'derived_from': [('level_id','=',self.level_id.id),('payscheme_id','=',self.payscheme_id.id)] }}
    
class ng_state_payroll_deduction_nonstd(models.Model):
    '''
    Non-Standard Deduction
    '''
    _name = "ng.state.payroll.deduction.nonstd"
    _description = 'Non-Standard Deduction'

    _columns = {
        'name': fields.char('Name', help='Deduction Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'permanent': fields.boolean('Permanent', help='Permanent'),
        'relief': fields.boolean('Relief', help='Forming part of CRF Relief'),
        'income_deduction': fields.boolean('Income Deduction', help='Deducted from Income before CRA 20% calculation'),
        'amount': fields.float('Amount', help='Amount', required=True),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'calendars': fields.many2many('ng.state.payroll.calendar', 'rel_deduction_nonstd_calendar', 'deduction_id','calendar_id', 'Calendars'),
    }

    _defaults = {
        'active': True,
        'permanent': False,
    }
    
    @api.onchange('name')
    def name_update(self):
        if self.name:
            self.name = self.name.upper()
	    self.name = self.name.strip()
    
class ng_state_payroll_earning_standard(models.Model):
    '''
    Standard Earning
    '''
    _name = "ng.state.payroll.earning.standard"
    _description = 'Standard Earning'

    _columns = {
        'name': fields.char('Name', help='Earning Name', required=True),
        'code': fields.char('Code', help='Rule Code', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'fixed': fields.boolean('Fixed Amount', help='Fixed Amount'),
        'taxable': fields.boolean('Taxable', help='Taxable'),
        'amount': fields.float('Amount', help='Amount', required=True),
        'payscheme_id': fields.many2one('ng.state.payroll.payscheme', 'Pay Scheme', required=False),
        'level_id': fields.many2one('ng.state.payroll.level', 'Grade', required=True),
        'derived_from': fields.many2one('ng.state.payroll.earning.standard', 'Derived From', required=False),
    }

    _defaults = {
        'active': True,
    }             
    
    @api.onchange('name')
    def name_update(self):
        if self.name:
            self.name = self.name.upper()
	    self.name = self.name.strip()
        
    @api.onchange('level_id','payscheme_id')
    def derived_from_id_update(self):
        return {'domain': {'derived_from': [('level_id','=',self.level_id.id),('payscheme_id','=',self.payscheme_id.id)] }}
        
class ng_state_payroll_earning_nonstd(models.Model):
    '''
    Non-Standard Earning
    '''
    _name = "ng.state.payroll.earning.nonstd"
    _description = 'Non-Standard Earning'

    _columns = {
        'name': fields.char('Name', help='Earning Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'permanent': fields.boolean('Permanent', help='Permanent'),
        'taxable': fields.boolean('Taxable', help='Taxable'),
        'amount': fields.float('Amount', help='Amount', required=True),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'calendars': fields.many2many('ng.state.payroll.calendar', 'rel_earning_nonstd_calendar', 'earning_id','calendar_id', 'Calendars'),
    }  

    _defaults = {
        'active': True,
        'permanent': False,
        'taxable': True,
    }
    
    @api.onchange('name')
    def name_update(self):
        if self.name:
            self.name = self.name.upper()
	    self.name = self.name.strip()
           
class ng_state_payroll_training(models.Model):
    '''
    Training
    '''
    _name = "ng.state.payroll.training"
    _description = 'Training'

    _columns = {
        'name': fields.char('Training', help='Training Name', required=True),
        'code': fields.char('Code', help='Training Code', required=True),
        'from_date': fields.date('From Date', help='Availability From Date', required=True),
        'to_date': fields.date('To Date', help='Availability To Date', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
    }  

    _defaults = {
        'active': True,
    }  
           
class ng_state_payroll_training_history(models.Model):
    '''
    Training History
    '''
    _name = "ng.state.payroll.training.history"
    _description = 'Training History'

    _columns = {
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'training_id': fields.many2one('ng.state.payroll.training', 'Training', required=True),
        'date': fields.date('Date Completed', help='Date Completed', required=False),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'state': fields.selection([
            ('unconfirmed', 'Unconfirmed'),
            ('confirmed', 'Confirmed'),
            ('fake', 'Not Genuine'),
            ], 'Verification Status', required=True, select=True),
    } 

    _defaults = {
        'active': True,
        'date': date.today(),
        'state': 'unconfirmed',
    } 
               
class ng_state_payroll_hrevent(models.Model):
    '''
    HR Activity
    '''
    _name = "ng.state.payroll.hrevent"
    _description = 'HR Activity'

    _columns = {
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'date': fields.datetime('Action Time', help='Exact time action occurred', required=True),
        'activity_type': fields.selection([
            ('retirement', 'Retirement'),
            ('transfer', 'Transfer'),
            ('changereq', 'Change Request'),
            ('promotion', 'Promotion'),
            ('suspension', 'Disciplinary - Suspension'),
            ('reinstatement', 'Disciplinary - Reinstatement'),
            ('loan', 'Loan'),
            ('demise', 'Demise'),
            ('termination', 'Termination'),
            ('query', 'Query'),
            ], 'Activity Type', required=True, select=True),
        'activity_id': fields.integer('Activity Reference', help='Activity Reference', required=True),

    } 

    _defaults = {
        'date': date.today(),
    }
    
    _rec_name = 'date'

    @api.multi
    def view_detail(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'ng.state.payroll.' + self.activity_type,
            'view_type': 'form',
            'view_mode': 'form',
            'res_id': self.activity_id,
            'target': 'new',
        }     
           
class ng_state_payroll_certification(models.Model):
    '''
    Certification
    '''
    _name = "ng.state.payroll.certification"
    _description = 'Certification'

    _columns = {
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'certificate_id': fields.many2one('ng.state.payroll.certificate', 'Certificate', required=True),
        'upload_id': fields.many2one('ng.state.payroll.certificate.upload', 'Upload', required=False),
        'date': fields.date('Date Awarded', help='Date Awarded', required=False),
        'expiration_date': fields.date('Expiration Date', help='Expiration Date', required=False),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'state': fields.selection([
            ('unconfirmed', 'Unconfirmed'),
            ('confirmed', 'Confirmed'),
            ('fake', 'Not Genuine'),
            ], 'Verification Status', required=True, select=True),
        'course_id': fields.many2one('ng.state.payroll.certcourse', 'Course of Study', required=False),

    } 

    _defaults = {
        'active': True,
        'date': date.today(),
        'expiration_date': date.today(),
        'state': 'unconfirmed',
    }
    
    _rec_name = 'employee_id'
           
class ng_state_payroll_certificate(models.Model):
    '''
    Certificate
    '''
    _name = "ng.state.payroll.certificate"
    _description = 'Certificate'

    _columns = {
        'name': fields.char('Certificate', help='Certificate Name', required=True),
        'code': fields.char('Code', help='Certificate Code', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'type': fields.selection([
            ('primary', 'Primary'),
            ('secondary', 'Secondary'),
            ('intermediate', 'Intermediate'),
            ('tertiary', 'Tertiary'),
            ('masters', 'Masters'),
            ('doctorate', 'Doctorate'),
            ('professional', 'Professional'),
            ('trade_test', 'Trade Test'),
            ], 'Certificate Type', required=True),
    } 

    _defaults = {
        'active': True,
    }  
           
class ng_state_payroll_certcourse(models.Model):
    '''
    Course of Study
    '''
    _name = "ng.state.payroll.certcourse"
    _description = 'Course of Study'

    _columns = {
        'name': fields.char('Certificate', help='Certificate Name', required=True),
        'code': fields.char('Code', help='Certificate Code', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
    } 

    _defaults = {
        'active': True,
    }  
           
class ng_state_payroll_stats(models.Model):
    '''
    Statistics
    '''
    _name = "ng.state.payroll.stats"
    _description = 'Statistics'

    _columns = {
        'name': fields.char('Attribute', help='Attribute Name', required=True),
        'count': fields.integer('Count', help='Count', required=True),
        'category': fields.selection([
            ('gender', 'Gender'),
            ('age', 'Age'),
            ('payscheme', 'Pay Scheme'),
            ('grade', 'Grade Level'),
            ('mda', 'MDA Category'),
            ], 'Statistical Category', required=True),
        'date': fields.datetime('Last Updated', help='Last Updated', required=False),
    }
    
    def try_init_stats(self, cr, uid, context=None):
        _logger.info("Running payroll cron-job...")
        cr.execute("update ng_state_payroll_stats set count=(select count(id) from hr_employee where gender='male' and (status_id=1 or status_id=2)) where category='gender' and name='Male'") 
        cr.execute("update ng_state_payroll_stats set count=(select count(id) from hr_employee where gender='female' and (status_id=1 or status_id=2)) where category='gender' and name='Female'")
        
        cr.execute("update ng_state_payroll_stats set count=(select count(id) from hr_employee where (extract(year from current_date) - extract(year from birthday)) >= 18 and (extract(year from current_date) - extract(year from birthday)) <= 30 and (status_id=1 or status_id=2)) where category='age' and name='18 - 30'")
        cr.execute("update ng_state_payroll_stats set count=(select count(id) from hr_employee where (extract(year from current_date) - extract(year from birthday)) >= 31 and (extract(year from current_date) - extract(year from birthday)) <= 40 and (status_id=1 or status_id=2)) where category='age' and name='31 - 40'")
        cr.execute("update ng_state_payroll_stats set count=(select count(id) from hr_employee where (extract(year from current_date) - extract(year from birthday)) >= 41 and (extract(year from current_date) - extract(year from birthday)) <= 50 and (status_id=1 or status_id=2)) where category='age' and name='41 - 50'")
        cr.execute("update ng_state_payroll_stats set count=(select count(id) from hr_employee where (extract(year from current_date) - extract(year from birthday)) >= 51 and (extract(year from current_date) - extract(year from birthday)) <= 60 and (status_id=1 or status_id=2)) where category='age' and name='51 - 60'")
        cr.execute("update ng_state_payroll_stats set count=(select count(id) from hr_employee where (extract(year from current_date) - extract(year from birthday)) >= 61 and (extract(year from current_date) - extract(year from birthday)) <= 80 and (status_id=1 or status_id=2)) where category='age' and name='61 - 80'")
        #TODO MDA Category = MDA, SUBEB, TESCOM, LGA, LTH
        
        cr.execute("update ng_state_payroll_stats set date=current_timestamp"); 
        
class ng_state_payroll_earning_employee(models.Model):
    '''
    Employee Earning
    '''
    _name = "ng.state.payroll.earning.employee"
    _description = 'Employee Earning'

    _columns = {
        'name': fields.char('Name', help='Earning Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'fixed': fields.boolean('Fixed Amount', help='Fixed Amount'),
        'taxable': fields.boolean('Taxable', help='Taxable'),
        'amount': fields.float('Amount', help='Amount', required=True),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'derived_from': fields.many2one('ng.state.payroll.earning.employee', 'Parent Earning'),
    }

    _defaults = {
        'active': True,
        'fixed': True,
    }   
            
class ng_state_payroll_deduction_employee(models.Model):
    '''
    Employee Deduction
    '''
    _name = "ng.state.payroll.deduction.employee"
    _description = 'Employee Deduction'

    _columns = {
        'name': fields.char('Name', help='Deduction Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'fixed': fields.boolean('Fixed Amount', help='Fixed Amount'),
        'amount': fields.float('Amount', help='Amount', required=True),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'derived_from': fields.many2one('ng.state.payroll.earning.employee', 'Parent Earning'),
        'bank_account_id': fields.many2one('res.partner.bank', 'Deduction Bank Account', required=True),
    }

    _defaults = {
        'active': True,
        'fixed': True,
    }             
        
class ng_state_payroll_subvention(models.Model):
    '''
    Subvention
    '''
    _name = "ng.state.payroll.subvention"
    _description = 'Subvention'

    _columns = {
        'name': fields.char('Name', help='Earning Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'amount': fields.float('Amount', help='Amount', required=True),
        'calendar_id': fields.many2one('ng.state.payroll.calendar', 'Calendar', required=True, track_visibility='onchange'),
        'org_id': fields.many2one('hr.department', 'MDA', required=True, select=True),
        'bank_id': fields.many2one('res.bank', string='Bank', required=True),
        'bank_account_no': fields.char('Bank Account', help='Bank Account Number', required=True),
    }

    _defaults = {
        'active': True,
    }             
       
class ng_state_payroll_salaryrule(models.Model):
    '''
    Salary Rule
    '''
    _name = "ng.state.payroll.salaryrule"
    _description = 'Salary Rule'

    _columns = {
        'code': fields.char('Code', help='Rule Code', required=True),
        'description': fields.char('Description', help='Rule Description', required=True),
    }
    
    _sql_constraints = [
        ('code_unique', 'unique(code)', 'Code already exists!')
    ]
    
class ng_state_payroll_calendar(models.Model):
    '''
    Pay Calendar
    '''
    _name = "ng.state.payroll.calendar"
    _description = 'Pay Calendar'

    _columns = {
        'name': fields.char('Name', help='Calendar Name', required=False),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'org_id': fields.many2one('hr.department', 'MDA', required=True, select=True),
        'schedule_pay': fields.selection([
            ('monthly', 'Monthly'),
            ('quarterly', 'Quarterly'),
            ('semi-annually', 'Semi-annually'),
            ('annually', 'Annually'),
            ('weekly', 'Weekly'),
            ('bi-weekly', 'Bi-weekly'),
            ('bi-monthly', 'Bi-monthly'),
            ], 'Scheduled Pay', required=True, select=True),
        'from_date': fields.date('From Date', help='From Date', required=True),
        'to_date': fields.date('To Date', help='To Date', required=True),
        'total_hours': fields.integer('Week Working Hours', required=False, help='Week Working Hours'),
    }

    _defaults = {
        'active': True,
        'total_hours': 40,
    }             

    def name_get(self, cr, uid, ids, context=None):
        if not ids:
            return []
        res = []
        for r in self.read(cr, uid, ids, ['id', 'from_date', 'to_date', 'name'], context):
            aux = ''
            if r['name']:
                aux = r['name']
    
            aux +=  " ("
            if r['from_date']:
                aux += datetime.strptime(r['from_date'], '%Y-%m-%d').strftime('%d/%m/%Y')
                # why translate a date? I think is a mistake, the _() function must have a 
                # known string, example _("the start date is %s") % r['from_date']
    
            aux +=  ' - '
            if r['to_date']:
                aux += datetime.strptime(r['to_date'], '%Y-%m-%d').strftime('%d/%m/%Y') # same question
    
            aux += ')'
    
            # aux is the name items for the r['id']
            res.append((r['id'], aux))  # append add the tuple (r['id'], aux) in the list res
    
        return res

        #Open create form with current month date range
    def name_search(self, cr, user, name='', args=None, operator='ilike', context=None, limit=100):
        if not args:
            args = []
        if name:
            ids = self.search(cr, user, [('name','=ilike',name)]+ args, limit=limit, context=context)
            if not ids:
                # Do not merge the 2 next lines into one single search, SQL search performance would be abysmal
                # on a database with thousands of matching products, due to the huge merge+unique needed for the
                # OR operator (and given the fact that the 'name' lookup results come from the ir.translation table
                # Performing a quick memory merge of ids in Python will give much better performance
                ids = set()
                ids.update(self.search(cr, user, args + [('to_date',operator,name)], limit=limit, context=context))
                if not limit or len(ids) < limit:
                    # we may underrun the limit because of dupes in the results, that's fine
                    ids.update(self.search(cr, user, args + [('from_date',operator,name)], limit=(limit and (limit-len(ids)) or False) , context=context))
                    #End
                ids = list(ids)
            if not ids:
                ptrn = re.compile('(\[(.*?)\])')
                res = ptrn.search(name)
                if res:
                    ids = self.search(cr, user, [('id','=', res.group(2))] + args, limit=limit, context=context)
        else:
            ids = self.search(cr, user, args, limit=limit, context=context)
        result = self.name_get(cr, user, ids, context=context)
        return result

class ng_state_payroll_taxrule(models.Model):
    '''
    Tax Rule
    '''
    _name = "ng.state.payroll.taxrule"
    _description = 'Tax Rule'

    _columns = {
        'name': fields.char('Name', help='Tax Rule Name', required=False),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'from_amount': fields.float('From Amount', help='From Amount', required=True),
        'to_amount': fields.float('To Amount', help='To Amount', required=True),
        'percentage': fields.float('Percentage', help='Percentage', required=True),
    }

    _defaults = {
        'active': True,
    }             

    def name_get(self, cr, uid, ids, context=None):
        if not ids:
            return []
        res = []
        for r in self.read(cr, uid, ids, ['id', 'name', 'from_amount', 'to_amount', 'percentage'], context):
            aux = ''
            if r['name']:
                aux = r['name']
                
            aux += '('
            if r['to_amount']:
                aux += str(r['to_amount'])

            aux += ' to '
            if r['to_amount']:
                aux += str(r['to_amount'])

            aux +=  ' @ '
            if r['percentage']:
                aux += (str(r['percentage']) + '%')

            aux += ')'
            # aux is the name items for the r['id']
            res.append((r['id'], aux))  # append add the tuple (r['id'], aux) in the list res
    
        return res

class ng_state_payroll_leaveallowance(models.Model):
    '''
    Leave Allowance
    '''
    _name = "ng.state.payroll.leaveallowance"
    _description = 'Leave Allowance'

    _columns = {
        'computation_base': fields.selection([
            ('basic', 'Basic'),
            ('basic_rent', 'Basic + Rent'),
        ], 'Computation Base'),
        'paygrade_id': fields.many2one('ng.state.payroll.paygrade', 'Pay Grade'),
        'percentage': fields.float('Percentage', help='Percentage', required=True),
        'payscheme_id': fields.many2one('ng.state.payroll.payscheme', 'Pay Scheme', required=True),
    }

    _defaults = {
        'computation_base': 'basic',
    }             

    def name_get(self, cr, uid, ids, context=None):
        if not ids:
            return []
        res = []
        for r in self.read(cr, uid, ids, ['id', 'paygrade_id', 'payscheme_id'], context):
            aux = ('(')
            if r['paygrade_id']:
                aux += ('Pay Grade - ' + str(r['paygrade_id'][1]) + ', ') # same question
    
            if r['payscheme_id']:
                aux += ('Pay Scheme - ' + r['payscheme_id'][1]) # same question
            aux += (')')
    
            # aux is the name items for the r['id']
            res.append((r['id'], aux))  # append add the tuple (r['id'], aux) in the list res
    
        return res
        
class ng_state_payroll_scenariobatch(models.Model):
    '''
    Scenario Batch
    '''
    _name = "ng.state.payroll.scenariobatch"
    _description = 'Scenario Batch'

    _columns = {
        'name': fields.char('Name', help='Scenario Name', required=True),
        'scenario_ids': fields.one2many('ng.state.payroll.scenario','batch_id','Scenarios'),
        'state': fields.selection([
            ('draft','Draft'),
            ('processed','Processed'),
            ('closed','Closed')
        ], 'Status')
    }
    
    @api.model
    def create(self, vals):
        vals['state'] = 'draft'
        res = super(ng_state_payroll_scenariobatch, self).create(vals)
            
        return res    
   
    @api.multi   
    def run_finalize(self):
        self.finalize()
   
    @api.multi   
    def run_dry_run(self):
        self.dry_run()

    @api.multi
    def dry_run(self, context=None):
        env = self.env
        with env.do_in_draft():
            res = self.finalize()
                   
        return res
         
    @api.multi
    def finalize(self):
        _logger.info("Calling finalize...state = %s", self.state)
        
        for scenario_id in self.scenario_ids:
            scenario_id.finalize()
        
        #Write processed if all scenarios completed successfully
        return self.write({'state':'processed'})  
    
class ng_state_payroll_scenario_signoff(models.Model):
    '''
    Payment Sign-Off
    '''
    _name = "ng.state.payroll.scenario.signoff"
    _description = 'Payment Sign-Off'

    _columns = {
        'scenario_id': fields.many2one('ng.state.payroll.scenario', 'Scenario', required=True),
        'user_id': fields.many2one('res.users', 'Payment Approver', required=True, domain="[('groups_id.name','=','Payroll Officer')]"),
        'signed_off': fields.boolean('Closed', help='Sign-Off closed status', required=True),
        'pos_order': fields.integer('Order', help='Order', required=True),
    }
    
    _defaults = {
        'signed_off': False,
    }                                         
    
class ng_state_payroll_stdconfig(models.Model):
    _name = "ng.state.payroll.stdconfig"
    _description = "Earnings & Deductions Configuration"

    _columns = {
        'name': fields.char('Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'initialized': fields.boolean('Initialized', help='Earnings & Deductions Initialized', required=True, readonly=1),
    } 
    
    _defaults = {
        'name': 'Basic Configuration',
        'initialized': True,
    } 
       
    def try_init_earn_dedt(self, cr, uid, context=None):
        _logger.info("Running try_init_earn_dedt: earnings/deductions cron-job...")
        stdconfig_singleton = self.pool.get('ng.state.payroll.stdconfig')
        stdconfig_ids = stdconfig_singleton.search(cr, uid, [('active', '=', True)], limit=1, context=context)
        if len(stdconfig_ids) == 1:
            stdconfig_obj = stdconfig_singleton.browse(cr, uid, stdconfig_ids[0], context=context)
            if not stdconfig_obj.initialized:
                _logger.info("Initializing earnings/deductions...")
                cr.execute("truncate rel_employee_std_earning")
                _logger.info("Truncated rel_employee_std_earning.")
                cr.execute("truncate rel_employee_std_deduction")
                _logger.info("Truncated rel_employee_std_deduction.")
                cr.execute("update hr_employee set resolved_earn_dedt='f'")
                _logger.info("Updated employee resolved_earn_dedt.")
                stdconfig_obj.init_earnings_deductions(context=context)
                stdconfig_obj.update({'initialized': True})
                _logger.info("Done initializing.")
        
        return True
                             
    @api.multi
    def init_earnings_deductions(self, context=None):
        employees = self.env['hr.employee'].search([('resolved_earn_dedt', '=', False), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id')
        _logger.info("init_earnings_deductions - Number of employees found: %d", len(employees))        
        
        tic = time.time()
        self.env.cr.execute('prepare insert_employee_std_earning (int, int) as insert into rel_employee_std_earning (employee_id,earning_id) values ($1, $2)')            
        self.env.cr.execute('prepare insert_employee_std_deduction (int, int) as insert into rel_employee_std_deduction (employee_id,deduction_id) values ($1, $2)')            
        for emp in employees:
            standard_earnings = self.env['ng.state.payroll.earning.standard'].search([('active', '=', True), ('payscheme_id', '=', emp.payscheme_id.id), ('level_id', '=', emp.level_id.id)])            
            for e in standard_earnings:
                self.env.cr.execute('execute insert_employee_std_earning(%s,%s)', (emp.id,e.id))
                
            standard_deductions = self.env['ng.state.payroll.deduction.standard'].search([('active', '=', True), ('payscheme_id', '=', emp.payscheme_id.id), ('level_id', '=', emp.level_id.id)])
            for d in standard_deductions:
                self.env.cr.execute('execute insert_employee_std_deduction(%s,%s)', (emp.id,d.id))
        if len(employees) > 0:
            self.env.cr.execute("update hr_employee set resolved_earn_dedt='t'")
            #self.env.cr.execute("update hr_employee set grade_level=(select level from ng_state_payroll_level where hr_employee.level_id=id) where resolved_earn_dedt='t'")
            #_logger.info("Updated grade levels.")
            self.env.cr.commit()
            _logger.info("Processed %d employees in %f seconds.", len(employees), (time.time() - tic))
    
    def try_resolve_earn_dedt(self, cr, uid, context=None):
        _logger.info("Running try_resolve_earn_dedt: earnings/deductions cron-job...")
        stdconfig_singleton = self.pool.get('ng.state.payroll.stdconfig')
        stdconfig_ids = stdconfig_singleton.search(cr, uid, [('active', '=', True)], limit=1, context=context)
        if len(stdconfig_ids) == 1:
            stdconfig_obj = stdconfig_singleton.browse(cr, uid, stdconfig_ids[0], context=context)
            stdconfig_obj.resolve_earnings_deductions(context=context)
        
        return True
                             
    @api.multi
    def resolve_earnings_deductions(self, context=None):
        _logger.info("resolve_earnings_deductions - updating resolved_earn_dedt flag...")
        self.env.cr.execute("update hr_employee set resolved_earn_dedt='f' where (status_id=1 or status_id=2) and id not in (select distinct employee_id from rel_employee_std_earning)")
        employees = self.env['hr.employee'].search([('resolved_earn_dedt', '=', False), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id')
        _logger.info("resolve_earnings_deductions - Number of employees found: %d", len(employees))        
        
        tic = time.time()            
        for emp in employees:
            self.env.cr.execute("delete from rel_employee_std_earning where employee_id=" + str(emp.id))
            _logger.info("Removed records for employee " + str(emp.name) + " [" + str(emp.id) + "] from Standard Earnings.")
            self.env.cr.execute("delete from rel_employee_std_deduction where employee_id=" + str(emp.id))
            _logger.info("Removed records for employee " + str(emp.name) + " [" + str(emp.id) + "] from Standard Deductions.")
            standard_earnings = self.env['ng.state.payroll.earning.standard'].search([('active', '=', True), ('payscheme_id', '=', emp.payscheme_id.id), ('level_id', '=', emp.level_id.id)])
            standard_deductions = self.env['ng.state.payroll.deduction.standard'].search([('active', '=', True), ('payscheme_id', '=', emp.payscheme_id.id), ('level_id', '=', emp.level_id.id)])
            
            for earning_std in standard_earnings:
                self.env.cr.execute("insert into rel_employee_std_earning (earning_id,employee_id) values (" + str(earning_std.id) + "," + str(emp.id) + ")")
            for deduction_std in standard_deductions:
                self.env.cr.execute("insert into rel_employee_std_deduction (deduction_id,employee_id) values (" + str(deduction_std.id) + "," + str(emp.id) + ")")
            self.env.cr.execute("update hr_employee set resolved_earn_dedt='t' where id=" + str(emp.id))
            _logger.info("Updated records for employee " + str(emp.name) + " [" + str(emp.id) + "].")
        #if len(employees) > 0:
            #self.env.cr.execute("update hr_employee set grade_level=(select level from ng_state_payroll_level where hr_employee.level_id=id) where resolved_earn_dedt='t'")
            #_logger.info("Updated grade levels.")
            #self.env.cr.commit()
            #_logger.info("Processed %d employees in %f seconds.", len(employees), (time.time() - tic))
                        
class ng_state_payroll_scenario(models.Model):
    '''
    Scenario
    '''
    _name = "ng.state.payroll.scenario"
    _description = 'Scenario'

    _columns = {
        'name': fields.char('Name', help='Scenario Name', required=True),
        'total_amount': fields.float('Total Payroll Paid Amount', help='Total Payroll Paid Amount'),
        'total_amount_pension': fields.float('Total Pension Paid Amount', help='Total Pension Paid Amount'),
        'processing_time': fields.float('Processing Time', help='Processing Time'),
        'batch_id': fields.many2one('ng.state.payroll.scenariobatch', 'Scenario Batch'),
        'payroll_id': fields.many2one('ng.state.payroll.payroll', 'Payroll', required=True),
        'scenario_item_ids': fields.one2many('ng.state.payroll.scenario.item','scenario_id','Payroll Scenario Items'),
        'scenario2_item_ids': fields.one2many('ng.state.payroll.scenario2.item','scenario_id','Pension Scenario Items'),
        'payment_ids': fields.one2many('ng.state.payroll.scenario.payment','scenario_id','Payroll Payment Items'),
        'payment2_ids': fields.one2many('ng.state.payroll.scenario2.payment','scenario_id','Pension Payment Items'),
        'signoff_ids': fields.one2many('ng.state.payroll.scenario.signoff','scenario_id','Sign-Off Items'),
        'signoff_pos_order': fields.integer('Sign-off Index', help='Sign-off Index'),
        'do_dry_run': fields.boolean('Dry Run', help='Tick check-box to do dry run'),
        'gov_sign': fields.binary(string='Governor Signature'),
        'ps_finance_sign': fields.binary(string='PS Finance Signature'),
        'employee_report': fields.binary('Employee Report'),
        'nibbs_report': fields.binary('NIBBS Report'),
        'exec_summary_report': fields.binary('Executive Summary Report'),
        'deduction_report': fields.binary('Deduction Report'),
        'mfb_report': fields.binary('MFB Report'),
        'mda_emails': fields.char('MDA Email', help='Comma separated email recipients for MDA notification', required=False),
        'generate_reports': fields.boolean('Generate Reports', help='Generate Reports'),
        'state': fields.selection([
            ('draft','Draft'),
            ('in_progress','Processing'),
            ('processed','Processed'),
            ('closed','Closed'),
            ('paid','Paid'),
        ], 'Status')
    }

    _defaults = {
        'state': 'draft',
        'signoff_pos_order': 0,
        'generate_reports':True,
    }
     
    @api.multi
    def reset_reports(self, vals):
        _logger.info("Calling reset_reports..vals = %s", vals)
        
        self.env.cr.execute("update ng_state_payroll_scenario set employee_report=null,nibbs_report=null,exec_summary_report=null,deduction_report=null where id=" + str(self.id))
        self.env.invalidate_all()
                
    @api.multi
    def copy(self, default=None):
        _logger.info("copy - %d, %s", self.id, default)            
            
        template = self.env['ng.state.payroll.scenario'].search([('id', '=', self.id)])
            
        scenario_item_ids = []
        scenario2_item_ids = []
        
        if not default.get('name'):
            default['name'] = _("%s (copy)") % (self.name)

        scenario = super(ng_state_payroll_scenario, self).copy(default)
        
        for item_id in template.scenario_item_ids:
            item_id_copy = {
                'scenario_id':scenario.id,
                'percentage':item_id.percentage,
                'level_min':item_id.level_min,
                'level_max':item_id.level_max,
                'payscheme_ids':item_id.payscheme_ids
            }
            scenario_item_ids.append(item_id_copy)
        
        for item_id in template.scenario2_item_ids:
            item_id_copy = {
                'scenario_id':scenario.id,
                'percentage':item_id.percentage,
                'amount_min':item_id.amount_min,
                'amount_max':item_id.amount_max
            }
            scenario2_item_ids.append(item_id_copy)
        
        scenario.update({
            'scenario_item_ids':scenario_item_ids,
            'scenario2_item_ids':scenario2_item_ids
        })

        return scenario
        
    @api.multi
    def revert(self):
        _logger.info("Calling revert..id = %s", self.id)
        #if self.env.user.has_group('hr_payroll_nigerian_state.group_payroll_admin'):
        if self.payroll_id.do_payroll:
            #self.env.cr.execute("with x as (select employee_id,balance_income,amount from ng_state_payroll_scenario_payment) update ng_state_payroll_payroll_item set balance_income = (x.balance_income + x.amount) from x where x.employee_id = ng_state_payroll_payroll_item.employee_id and ng_state_payroll_payroll_item.payroll_id=" + str(self.id))
            for p in self.payment_ids:
                p.payroll_item_id.write({'balance_income': (p.payroll_item_id.balance_income + p.amount)})
            self.env.cr.execute("delete from ng_state_payroll_scenario_payment where scenario_id=" + str(self.id))

        if self.payroll_id.do_pension:
            #self.env.cr.execute("with x as (select employee_id,balance_income,amount from ng_state_payroll_scenario2_payment) update ng_state_payroll_pension_item set balance_income = (x.balance_income + x.amount) from x where x.employee_id = ng_state_payroll_pension_item.employee_id and ng_state_payroll_payroll_item.payroll_id=" + str(self.id))        
            for p in self.payment2_ids:
                p.pension_item_id.write({'balance_income': (p.pension_item_id.balance_income + p.amount)})
            self.env.cr.execute("delete from ng_state_payroll_scenario2_payment where scenario_id=" + str(self.id))

        self.env.cr.execute("update ng_state_payroll_scenario set total_amount=0,processing_time=0,state='draft'")
        self.env.cr.execute("delete from ng_state_payroll_scenario_signoff where scenario_id=" + str(self.id))
        self.env.cr.execute("update ng_state_payroll_scenario set employee_report=null,nibbs_report=null,exec_summary_report=null,deduction_report=null where id=" + str(self.id))
        self.env.invalidate_all()
                                                                 
    #On create; iterate through levels and create new scenario items
    #Method to do dry run
    #Method to save
    @api.model
    def create(self, vals):
        vals['state'] = 'draft'
        res = super(ng_state_payroll_scenario, self).create(vals)
            
        return res

    @api.multi
    def write(self, vals):
        _logger.info("Writing scenario..", vals)
        
        if ('do_dry_run' in vals and vals['do_dry_run']) and vals['state'] == 'draft':
            vals['state'] = 'processed'
        
        return super(ng_state_payroll_scenario,self).write(vals)
   
    @api.multi   
    def run_finalize(self):
        return self.finalize()
   
    @api.multi   
    def run_dry_run(self):
        return self.dry_run()
        
    @api.multi
    def sign_off(self):        
        _logger.info("Calling sign_off..state = %s", self.state)
        #TODO Set sign-off entry for current user to true
        group_payroll_officer = self.env['res.groups'].search([('name', '=', 'Payroll Officer')])
        group_admin = self.env['res.groups'].search([('name', '=', 'Configuration')])
        #if group_payroll_officer in self.env.user.groups_id or group_admin in self.env.user.groups_id:
        if True:
            #Iterate through sign-off users and if all signed off, set state='closed'
            signoff_count = 0
            for sign_off in self.signoff_ids:
                if sign_off.user_id.id == self.env.user.id:
                    self.update({'signoff_pos_order': (self.signoff_pos_order + 1)})
                    sign_off.update({'signed_off': True})
                if sign_off.signed_off:
                    signoff_count += 1
            if len(self.signoff_ids) == signoff_count:
                self.state = 'closed'
                self.update({'state': 'closed'})        
   
    @api.multi
    def set_in_progress(self):
        self.write({'state': 'in_progress'})
        
    def dry_run(self):
        _logger.info("Calling dry_run...state = %s", self.state)        
        if self.state == 'in_progress':
            raise osv.except_osv(_('Info'), _('Processing already in progress.'))

        if not self.state == 'in_progress':        
            self.set_in_progress()            
            #Payment for payroll
            payroll_items = self.payroll_id.payroll_item_ids
            total_amount = 0
            payment_item_list = []        
            if self.payroll_id.total_balance_payroll > 0:
                for payroll_item in payroll_items:
                    if payroll_item.balance_income > 0:
                        scenario_item = False
                        for s_item in self.scenario_item_ids:
                            if payroll_item.employee_id.payscheme_id.id in s_item.payscheme_ids.ids and payroll_item.employee_id.level_id.paygrade_id.level >= s_item.level_min and payroll_item.employee_id.level_id.paygrade_id.level <= s_item.level_max:
                                scenario_item = s_item
                                break
                        if scenario_item:
                            #Calculate the amount to be paid as a percentage of the Net
                            #If the amount is greater than the balance, pay the entire balance
                            amount = scenario_item.percentage * payroll_item.net_income / 100
                            if amount > payroll_item.balance_income:
                                amount = payroll_item.balance_income
                            total_amount += amount
                            payment_item = {
                                'employee_id': payroll_item.employee_id.id,
                                'active': True,
                                'amount': amount,
                                'payroll_item_id': payroll_item.id,
                                'balance_income': payroll_item.balance_income - amount,
                                'net_income': payroll_item.net_income,
                                'percentage': scenario_item.percentage,
                                'scenario_id': self.id
                            }
                            payment_item_list.append(payment_item)
                            payroll_item.update({'balance_income':payroll_item.balance_income - amount})
                self.total_amount = total_amount
                self.payment_ids = payment_item_list
    
            #Payment for pension
            pension_items = self.payroll_id.pension_item_ids
            total_amount = 0
            payment_item_list = []        
            if self.payroll_id.total_balance_pension > 0:
                for pension_item in pension_items:
                    if pension_item.balance_income > 0:
                        scenario2_item = False
                        for s_item in self.scenario2_item_ids:
                            if (pension_item.employee_id.annual_pension / 12) >= s_item.amount_min and (payroll_item.employee_id.annual_pension / 12) <= s_item.amount_max:
                                scenario2_item = s_item
                        if scenario2_item:
                            #Calculate the amount to be paid as a percentage of the Net
                            #If the amount is greater than the balance, pay the entire balance
                            amount = scenario2_item.percentage * (pension_item.employee_id.annual_pension / 12) / 100
                            if amount > pension_item.balance_income:
                                amount = pension_item.balance_income
                            total_amount += amount
                            payment_item = {
                                'employee_id': pension_item.employee_id.id,
                                'active': True,
                                'amount': amount,
                                'pension_item_id': pension_item.id,
                                'balance_income': pension_item.balance_income - amount,
                                'net_income': pension_item.net_income,
                                'percentage': scenario_item.percentage,
                                'scenario_id': self.id
                            }
                            payment_item_list.append(payment_item)
                            pension_item.update({'balance_income':payment_item.balance_income - amount})
                self.total_amount_pension = total_amount
                self.payment2_ids = payment_item_list

    @api.multi
    def finalize(self):
        _logger.info("Calling finalize...state = %s", self.state)
        if self.state == 'in_progress':
            raise osv.except_osv(_('Info'), _('Processing already in progress.'))
        
        if not self.state == 'in_progress':        
            self.set_in_progress()             
    
            #Payment for payroll
            payroll_items = self.payroll_id.payroll_item_ids
            total_amount = 0
            amount = 0
            if self.payroll_id.total_balance_payroll > 0:
                for payroll_item in payroll_items:
                    if payroll_item.balance_income > 0 and payroll_item.active:
                        scenario_item = False
                        for s_item in self.scenario_item_ids:
                            if payroll_item.employee_id.payscheme_id in s_item.payscheme_ids and payroll_item.employee_id.level_id.paygrade_id.level >= s_item.level_min and payroll_item.employee_id.level_id.paygrade_id.level <= s_item.level_max:
                                scenario_item = s_item
                        if scenario_item:
                            #Calculate the amount to be paid as a percentage of the Net
                            #If the amount is greater than the balance, pay the entire balance
                            amount = scenario_item.percentage * payroll_item.net_income / 100
                            if amount > payroll_item.balance_income:
                                amount = payroll_item.balance_income
                            total_amount += amount
                            payment_item = {'employee_id': payroll_item.employee_id.id,
                                                'active': True,
                                                'amount': amount,
                                                'payroll_item_id': payroll_item.id,
                                                'balance_income': payroll_item.balance_income - amount,
                                                'net_income': payroll_item.net_income,
                                                'percentage': scenario_item.percentage,
                                                'scenario_id': self.id}
                            self.env['ng.state.payroll.scenario.payment'].create(payment_item)
                            payroll_item.write({'balance_income':payroll_item.balance_income - amount})
                self.payroll_id.write({'total_balance_payroll': self.payroll_id.total_balance_payroll - amount})
                self.write({'state':'processed','total_amount':total_amount})
    
            #Payment for pension
            pension_items = self.payroll_id.pension_item_ids
            total_amount = 0
            if self.payroll_id.total_balance_pension > 0:
                for pension_item in pension_items:
                    if pension_item.balance_income > 0:
                        scenario2_item = False
                        for s_item in self.scenario2_item_ids:
                            if (pension_item.employee_id.annual_pension / 12) >= s_item.amount_min and (pension_item.employee_id.annual_pension / 12) <= s_item.amount_max:
                                scenario2_item = s_item
                        if scenario2_item:
                            #Calculate the amount to be paid as a percentage of the Net
                            #If the amount is greater than the balance, pay the entire balance
                            amount = scenario2_item.percentage * (pension_item.employee_id.annual_pension / 12) / 100
                            if amount > pension_item.balance_income:
                                amount = pension_item.balance_income
                            total_amount += amount
                            payment_item = {'employee_id': pension_item.employee_id.id,
                                                'active': True,
                                                'amount': amount,
                                                'pension_item_id': pension_item.id,
                                                'balance_income': (pension_item.balance_income - amount),
                                                'net_income': pension_item.net_income,
                                                'percentage': scenario2_item.percentage,
                                                'scenario_id': self.id}
                            self.env['ng.state.payroll.scenario2.payment'].create(payment_item)
                            pension_item.write({'balance_income':(pension_item.balance_income - amount)})
                self.payroll_id.write({'total_balance_pension': (self.payroll_id.total_balance_pension - amount)})
                self.write({'state':'processed','total_amount_pension':total_amount})            
     
    def try_generate_reports(self, cr, uid, context=None):
        _logger.info("Running try_generate_reports cron-job...")
        scenario_singleton = self.pool.get('ng.state.payroll.scenario')
        scenario_ids = scenario_singleton.search(cr, uid, [('state', '=', 'closed'), ('generate_reports', '=', True)], context=context)
        scenario_obj = None
        for scenario_id in scenario_ids:
            scenario_obj = scenario_singleton.browse(cr, uid, scenario_id, context=context)
            scenario_obj.process_reports()

        return True
       
    @api.multi
    def process_reports(self):
        _logger.info("process_reports : %s", self.mda_emails)
        if self.mda_emails:
            path = '/tmp/scenario_reports_' + str(self.id)
            if not os.path.exists(path):
                os.makedirs(path)
    
            _logger.info("process_reports : payment_item_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payment_item_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/payment_item_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : payment_nibbs_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payment_nibbs_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/payment_nibbs_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : payment_exec_summary_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payment_exec_summary_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/payment_exec_summary_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : deduction_nibbs_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = deduction_nibbs_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/deduction_nibbs_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : payment_mfb_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payment_mfb_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/payment_mfb_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            shutil.make_archive(path, 'zip', path)

            receivers = self.mda_emails #Comma separated email addresses
            message = "Dear Sir,\nPlease find the reports for the finalized scenario as found in the title of this email.\n\nThank you.\n"
            smtp_obj = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
            smtp_obj.ehlo()
            #smtp_obj.starttls()
            #smtp_obj.ehlo()
            smtp_obj.login(user="osun.payroll@gmail.com", password="p@55w0rd1939")
            sender = 'osun.payroll@gmail.com'
            msg = MIMEMultipart()
            msg['Subject'] = 'Scenario Closed - ' + self.name 
            msg['From'] = sender
            #msg['To'] = ', '.join(receivers)
            msg['To'] = receivers
                             
            part = MIMEBase('application', "octet-stream")
            part.set_payload(open(path + '.zip', "rb").read())
            Encoders.encode_base64(part)                            
            part.add_header('Content-Disposition', 'attachment; filename="scenario_reports_' + str(self.id) + '.zip"')
            msg.attach(MIMEText(message))
            msg.attach(part)
            smtp_obj.sendmail(sender, receivers, msg.as_string())
            self.write({'generate_reports':False})           
                
class ng_state_payroll_scenario_item(models.Model):
    '''
    Payroll Scenario Item
    '''
    _name = "ng.state.payroll.scenario.item"
    _description = 'Payroll Scenario Item'

    _columns = {
        'name': fields.char('Name', required=True),
        'percentage': fields.float('Percentage', help='Percentage', default=100, required=True),
        'level_min': fields.integer('Minimum Grade Level', help='Minimum Grade Level', required=True),
        'level_max': fields.integer('Maximum Grade Level', help='Maximum Grade Level'),
        'scenario_id': fields.many2one('ng.state.payroll.scenario', 'Scenario', required=True),
        'payscheme_ids': fields.many2many('ng.state.payroll.payscheme', 'rel_scenarioitem_payscheme', 'item_id','payscheme_id', 'Pay Schemes'),
    }
     
    _defaults = {
        'level_max': 100000,
    }
    
class ng_state_payroll_scenario2_item(models.Model):
    '''
    Pension Scenario Item
    '''
    _name = "ng.state.payroll.scenario2.item"
    _description = 'Pension Scenario Item'

    _columns = {
        'name': fields.char('Name', required=True),
        'percentage': fields.float('Percentage', help='Percentage', default=100, required=True),
        'amount_min': fields.float('Minimum Amount', help='Minimum Amount', required=True),
        'amount_max': fields.float('Maximum Amount', help='Maximum Amount'),
        'scenario_id': fields.many2one('ng.state.payroll.scenario', 'Scenario', required=True),
    }
     
    _defaults = {
        'amount_max': 1000000000,
    }
    
class ng_state_payroll_payroll_schoolsummary(models.Model):
    '''
    Summary Item
    '''
    _name = "ng.state.payroll.payroll.schoolsummary"
    _description = 'Payroll School Summary Item'

    _columns = {
        'school_id': fields.many2one('ng.state.payroll.school', 'School', required=True),
        'school': fields.related('school_id', 'name', type='char', string='School Name', readonly=1),
        'payroll_id': fields.many2one('ng.state.payroll.payroll', 'Payroll'),
        'total_taxable_income': fields.float('Taxable', help='Total Taxable Income'),
        'total_gross_income': fields.float('Actual Gross', help='Actual Total Gross'),
        'total_gross_expected': fields.float('Expected Gross', help='Expected Total Gross'),
        'total_net_income': fields.float('Net', help='Total Net'),
        'total_paye_tax': fields.float('Tax', help='Total PAYE Tax'),
        'total_nhf': fields.float('Tax', help='Total NHF'),
        'total_pension': fields.float('Pension', help='Total Contributory Pension'),
        'total_other_deductions': fields.float('Other Deductions', help='Total Other Deductions'),
        'total_leave_allowance': fields.float('Leave All.', help='Leave Allowance'),
        'total_strength': fields.integer('Staff Strength', help='Total Staff Strength'),
        'payslips_zip': fields.binary('Group Payslips'),
    }
     
    _defaults = {
        'total_gross_income': 0,
        'total_gross_expected': 0,
        'total_taxable_income': 0,
        'total_net_income': 0,
        'total_paye_tax': 0,
        'total_nhf': 0,
        'total_pension': 0,
        'total_other_deductions': 0,
        'total_leave_allowance': 0,
        'total_strength': 0,
    }
    
class ng_state_payroll_payroll_summary(models.Model):
    '''
    Summary Item
    '''
    _name = "ng.state.payroll.payroll.summary"
    _description = 'Payroll Summary Item'

    _columns = {
        'department_id': fields.many2one('hr.department', 'MDA', required=True),
        'department': fields.related('department_id', 'name', type='char', string='MDA Name', readonly=1),
        'payroll_id': fields.many2one('ng.state.payroll.payroll', 'Payroll'),
        'total_taxable_income': fields.float('Taxable', help='Total Taxable Income'),
        'total_gross_income': fields.float('Actual Gross', help='Actual Total Gross'),
        'total_gross_expected': fields.float('Expected Gross', help='Expected Total Gross'),
        'total_net_income': fields.float('Net', help='Total Net'),
        'total_paye_tax': fields.float('Tax', help='Total PAYE Tax'),
        'total_nhf': fields.float('Tax', help='Total NHF'),
        'total_pension': fields.float('Pension', help='Total Contributory Pension'),
        'total_other_deductions': fields.float('Other Deductions', help='Total Other Deductions'),
        'total_leave_allowance': fields.float('Leave All.', help='Leave Allowance'),
        'total_strength': fields.integer('Staff Strength', help='Total Staff Strength'),
        'payslips_zip': fields.binary('Group Payslips'),
    }
     
    _defaults = {
        'total_gross_income': 0,
        'total_gross_expected': 0,
        'total_taxable_income': 0,
        'total_net_income': 0,
        'total_paye_tax': 0,
        'total_nhf': 0,
        'total_pension': 0,
        'total_other_deductions': 0,
        'total_leave_allowance': 0,
        'total_strength': 0,
    }
        
class ng_state_payroll_pension_summary(models.Model):
    '''
    Summary Item
    '''
    _name = "ng.state.payroll.pension.summary"
    _description = 'Pension Summary Item'

    _columns = {
        'tco_id': fields.many2one('ng.state.payroll.tco', 'TCO', required=True),
        'tco': fields.related('tco_id', 'name', type='char', string='TCO Name', readonly=1),
        'total_gross_income': fields.float('Gross', help='Total Gross'),
        'total_gross_expected': fields.float('Expected Gross', help='Expected Total Gross'),
        'total_net_income': fields.float('Net', help='Total Net'),
        'total_arrears': fields.float('Arrears', help='Total Arrears'),
        'total_dues': fields.float('Dues', help='Total Dues'),
        'total_strength': fields.integer('Staff Strength', help='Total Staff Strength'),
        'payroll_id': fields.many2one('ng.state.payroll.payroll', 'Payroll'),
    }

    _defaults = {
        'total_gross_income': 0,
        'total_gross_expected': 0,
        'total_net_income': 0,
        'total_arrears': 0,
        'total_dues': 0,
        'total_strength': 0,
    }                                   

class ng_state_payroll_subvention_item(models.Model):
    '''
    Subvention Item
    '''
    _name = "ng.state.payroll.subvention.item"
    _description = 'Item'

    _columns = {
        'department_id': fields.many2one('hr.department', 'MDA', required=True),
        'name': fields.char('Name', help='Earning Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'amount': fields.float('Amount', help='Amount', required=True),
        'payroll_id': fields.many2one('ng.state.payroll.payroll', 'Payroll'),
    }
    
    _defaults = {
        'amount': 0,
        'active': True,
    }                                          
    
class ng_state_payroll_chamsswitch_config(models.Model):
    _name = "ng.state.payroll.chamsswitch.config"
    _description = "Chams Switch Configuration"

    _columns = {
        'name': fields.char('Name', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'testing': fields.boolean('Testing', help='Testing', required=False),
        'test_content': fields.text('Test Content', required=True),
        'url': fields.char('Chams Switch URL', required=True),
        'http_meth': fields.selection([
            ('post', 'POST'), 
            ('get', 'GET'), 
            ('put', 'PUT'), 
            ('patch', 'PATCH'), 
            ('delete', 'DELETE')], string='HTTP Method'),
        'user': fields.char('Username', required=False),
        'passwd': fields.char('Password', required=False),
    } 
    
    _defaults = {
        'name': 'Test Configuration',
        'active': True,
        'testing': True,
        'url': 'http://e2e6dbbd.ngrok.io/ChamsPay/rest/chams/upload',
        'http_meth': 'post',
        'test_content': '{"terminalId":"0000001","totalAmount":10000,"totalCount":3,"scheduleIdentifier":"jfduuf020000","uploadList":[{"staffId":"1","name":"James Doe","emailAddress":"doe@gmail.com","mobile":"080","amount":1000,"bankCode":"011","accountNumber":"1122334455","paymentId":"0001"},{"staffId":"2","name":"Jane Doe","emailAddress":"jane.doe@gmail.com","mobile":"080","amount":2000,"bankCode":"014","accountNumber":"5544332211","paymentId":"0002"},{"staffId":"3","name":"Shane Doe","emailAddress":"shane.doe@gmail.com","mobile":"080","amount":3000,"bankCode":"032","accountNumber":"5544332211","paymentId":"0003"}]}'
    } 
    
class ng_state_payroll_chamsswitch_batch(models.Model):
    '''
    Chams Switch Batch Payment
    '''
    _name = "ng.state.payroll.chamsswitch.batch"
    _description = 'Chams Switch Batch Payment'
     
    _columns = {
        'name': fields.char('Batch Name (Narration)', help='Batch Name (Narration)', required=True),
        'scenario_id': fields.many2one('ng.state.payroll.scenario', 'Scenario', domain="[('state', '=', 'closed')]", required=True),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('approved', 'Approved'),
            ('processed', 'Processed'),
            ('sent', 'Sent'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'payment_ids': fields.one2many('ng.state.payroll.chamsswitch.payment','batch_id','Payment Lines'),
    }      
    
    _defaults = {
        'state': 'draft',
    }
     
    @api.multi
    def cancel(self):
        _logger.info("cancel - %s", 'cancel')
        self.write({'state':'cancel'})
     
    @api.multi
    def confirm(self):
        _logger.info("confirm - %s", 'confirm')
        self.write({'state':'confirm'})
            
    @api.multi
    def approve(self):
        _logger.info("approve - %s", 'approve')
        self.write({'state':'approved'})
        
    @api.model
    def update_payment_status(self, schedule_id, staff_id, status):
        _logger.info("Calling update_payment_status...%s, %s", schedule_id, staff_id, status)
        
        if schedule_id and staff_id and status:
            payment_instance = self.env['ng.state.payroll.chamsswitch.payment'].search([('employee_id', '=', staff_id),('batch_id', '=', schedule_id)])
            if payment_instance:
                payment_instance.write({'state': status})
            else:
                return "No payment instance with payment_id '" + staff_id + "' found."
        else:
            return "Request has no payment_id and/or status parameters."
        
        return True
        
    @api.model
    def update_payment_statuses(self, schedule_id, serial_nos, statuses):
        _logger.info("Calling update_payment_status...%s, %s, %s", schedule_id, serial_nos, statuses)
        
        if schedule_id and serial_nos and statuses:
            if len(serial_nos) != len(statuses):
                return "Number of payment_ids must match number of statuses"
            
            for idx in range(len(serial_nos)):
                
                payment_instance = self.env['ng.state.payroll.chamsswitch.payment'].search([('serial_no', '=', serial_nos[idx]),('batch_id', '=', int(schedule_id[4:]))])
                if payment_instance:
                    payment_instance.write({'state': statuses[idx]})
                else:
                    return "No payment instance with payment_id '" + str(serial_nos[idx]) + "' found or schedule with ID " + schedule_id + " found."
        else:
            return "Request has no payment_id and/or status parameters."
        
        return True
    
    @api.multi
    def process_payment(self, context=None):
        _logger.info("Calling process_payment...state = %s", self.state)
        for payment_obj in self:
            _logger.info("payment_obj...state = %s", payment_obj.state)
            if payment_obj.state == 'approved':
                serial_no = 1
                
                payment_ids = payment_obj.scenario_id.payment_ids.filtered(lambda r: r.active == True)
                payments = []
                if payment_ids:
                    for item in payment_ids:
                        
                        payments.append({'serial_no':serial_no,'batch_id':payment_obj.id,'employee_id':item.employee_id.id,'amount':item.amount})
                        serial_no += 1
    
                payment2_ids = payment_obj.scenario_id.payment2_ids.filtered(lambda r: r.active == True)
                if payment2_ids:
                    for item in payment2_ids:
                        payments.append({'serial_no':serial_no,'batch_id':payment_obj.id,'employee_id':item.employee_id.id,'amount':item.amount})
                        serial_no += 1
                
                payment_obj.write({'state':'processed', 'payment_ids':[(0, 0, x) for x in payments]})
    
    @api.multi   
    def send_batch(self, context=None):
        _logger.info("Calling send_batch...")
        
        config_instance = self.env['ng.state.payroll.chamsswitch.config'].search([('active', '=', True)])
        active_config = config_instance[0]
        auth = None
        if active_config.user and active_config.passwd:
            auth = (active_config.user, active_config.passwd)
        headers = {'Content-Type': 'application/json', 'Accept': 'application/json'}
        
        req = active_config.test_content
        if not active_config.testing and self:
            for payment_obj in self:            
                #TODO Generate JSON request
                req = {
                    'terminalId':'0000001',
                    'totalAmount':0,
                    'totalCount':len(payment_obj.payment_ids),
                    'scheduleIdentifier':'OSSG' + str(self.id).zfill(6),
                    'uploadList':[]
                }
                totalAmount = 0
                uploadList = []
                for payment_id in payment_obj.payment_ids:
                    totalAmount += payment_id.amount
                    uploadList.append({
                        'paymentId':payment_id.serial_no,
                        'staffId':payment_id.employee_id.id,
                        'name':payment_id.employee_id.name_related,
                        'emailAddress':payment_id.employee_id.work_email,
                        'mobile':payment_id.employee_id.mobile_phone,
                        'accountNumber':payment_id.employee_id.bank_account_no,
                        'bankCode':payment_id.employee_id.bank_id.code,
                        'amount':payment_id.amount
                    })
                payment_obj.write({'state':'sent'})
                req.update({'totalAmount':totalAmount,'uploadList':uploadList})
                result = getattr(requests, active_config.http_meth)(active_config.url, json.dumps(req), auth=auth, headers=headers)
                _logger.info("send_batch result=%s", result.text)
            
        return True
    
class ng_state_payroll_chamsswitch_payment(models.Model):
    '''
    Chams Switch Payment
    '''
    _name = "ng.state.payroll.chamsswitch.payment"
    _description = 'Chams Switch Payment'
    
    _columns = {
        'batch_id': fields.many2one('ng.state.payroll.chamsswitch.batch', 'Payment Batch', required=True),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'serial_no': fields.char('Serial Number', help='Serial Number', required=True),
        'amount': fields.float('Amount', help='Amount', required=True),
        'state': fields.selection([
            ('00', 'Successful'),
            ('01', 'DUPLICATE UPLOAD'),
            ('02', 'MANDATORY FIELD NOT SET'),
            ('03', 'UNKNOWN TERMINAL ID'),
            ('05', 'FORMAT ERROR'),
            ('06', 'IN PROGRESS'),
            ('99', 'SYSTEM ERROR'),
            ('09', 'REJECTED BY APPROVER'),
            ('10', 'INVALID NUBAN NUMBER'),
            ('11', 'PAYMENT DISHONOURED BY BANK'),
            ('12', 'PROCESSING COMPLETED WITH ERROR'),
            ('14', 'RECORD NOT FOUND'),
            ('15', 'PAYMENT FAILED'),
            ('16', 'REQUEST ACCEPTED'),
            ('17', 'PAYMENT SCHEDULE NOT FOUND'),
            ('18', 'INVALID SCHEDULE ID'),
            ('19', 'BENEFICIARY BANK NOT AVAILABLE'),
            ('20', 'DO NOT HONOR'),
            ('21', 'DORMANT ACCOUNT'),
            ('22', 'INVALID BANK CODE'),
            ('23', 'INVALID BANK ACCOUNT'),
            ('24', 'CANNOT VERIFY ACCOUNT'),
        ], 'State', required=False, readonly=True),
    }
    
class ng_state_payroll_scenario_payment(models.Model):
    '''
    Payroll Payment Item
    '''
    _name = "ng.state.payroll.scenario.payment"
    _description = 'Payment Item'

    _columns = {
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'net_income': fields.float('Net Income', help='Net Income for calendar period', required=True),
        'balance_income': fields.float('Payment Balance', help='Balance of paid income for calendar period', required=True),
        'amount': fields.float('Paid Amount', help='Amount paid out of expected Net', required=True),
        'percentage': fields.float('Percentage', help='Percentage of Net Salary Paid', required=True),
        'payroll_item_id': fields.many2one('ng.state.payroll.payroll.item', 'Payroll Item', required=True),
        'scenario_id': fields.many2one('ng.state.payroll.scenario', 'Scenario', required=True), 
    }
    
    _defaults = {
        'amount': 0,
        'percentage': 0,
        'active': True,
    }
        
    @api.multi
    def update_payment_status(self, payment_id, status, context=None):
        _logger.info("Calling update_payment_status...%d, %s", payment_id, status)
        
        payment_instance = self.env['ng.state.payroll.scenario.payment'].search([('id', '=', payment_id)])
        
        payment_instance.write({'state': status})
        
        return True                                            

class ng_state_payroll_scenario2_payment(models.Model):
    '''
    Payment Item
    '''
    _name = "ng.state.payroll.scenario2.payment"
    _description = 'Pension Payment Item'

    _columns = {
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'net_income': fields.float('Net Income', help='Net Income for calendar period', required=True),
        'balance_income': fields.float('Payment Balance', help='Balance of paid income for calendar period', required=True),
        'amount': fields.float('Paid Amount', help='Amount paid out of expected Net', required=True),
        'percentage': fields.float('Percentage', help='Percentage of Net Salary Paid', required=True),
        'pension_item_id': fields.many2one('ng.state.payroll.pension.item', 'Pension Item', required=True),
        'scenario_id': fields.many2one('ng.state.payroll.scenario', 'Scenario', required=True),        
    }
    
    _defaults = {
        'amount': 0,
        'percentage': 0,
        'active': True,
    }                                                                                   

class ng_state_payroll_pension_item(models.Model):
    '''
    Pension Item
    '''
    _name = "ng.state.payroll.pension.item"
    _description = 'Pension Item'

    _columns = {
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'payment_item_ids': fields.one2many('ng.state.payroll.scenario2.payment','scenario_id','Payment Items', compute='_compute_payment_items'),
        'gross_income': fields.float('Gross', help='Gross Income'),
        'net_income': fields.float('Net', help='Net Income'),
        'balance_income': fields.float('Unpaid', help='Unpaid Balance'),
        'arrears_amount': fields.float('Arrears', help='Total arrears'),
        'payroll_id': fields.many2one('ng.state.payroll.payroll', 'Payroll'),
        'tco_id': fields.many2one('ng.state.payroll.tco', 'TCO', required=True),
        'item_line_ids': fields.one2many('ng.state.payroll.pension.item.line','item_id','Pension Item Lines'),
    }
    
    _defaults = {
        'gross_income': 0,
        'net_income': 0,
        'active': True,
    }        
   
    @api.depends('payroll_id', 'employee_id')   
    def _compute_payment_items(self):
        for payroll_item in self:
            payroll_item.payment_item_ids = self.env['ng.state.payroll.scenario2.payment'].search([('employee_id.id', '=', payroll_item.employee_id.id), ('scenario_id.payroll_id.id', '=', payroll_item.payroll_id.id)])        
    
class ng_state_payroll_payroll_item(models.Model):
    '''
    Payroll Item
    '''
    _name = "ng.state.payroll.payroll.item"
    _description = 'Payroll Item'

    _columns = {
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True),
        'active': fields.boolean('Active', help='Active Status', required=True),
        'resolve': fields.boolean('Resolve', help='Requires Resolution'),
        'retiring': fields.boolean('Retiring', help='Retiring this calendar period'),
        'payment_item_ids': fields.one2many('ng.state.payroll.scenario.payment','scenario_id','Payment Items', compute='_compute_payment_items'),
        'item_line_ids': fields.one2many('ng.state.payroll.payroll.item.line','item_id','Payroll Item Lines'),
        'taxable_income': fields.float('Taxable', help='Taxable Income'),
        'gross_income': fields.float('Gross', help='Gross Income'),
        'net_income': fields.float('Net', help='Net Income'),
        'leave_allowance': fields.float('Leave Allowance', help='Leave Allowance'),
        'balance_income': fields.float('Unpaid', help='Unpaid Balance'),
        'paye_tax': fields.float('Monthly Tax', help='Monthly PAYE Tax'),
        'paye_tax_annual': fields.float('Annual Tax', help='Annual PAYE Tax'),
        'payroll_id': fields.many2one('ng.state.payroll.payroll', 'Payroll'),
        'paycategory_id': fields.many2one('ng.state.payroll.paycategory', 'Pay Category'),
        'payscheme_id': fields.many2one('ng.state.payroll.payscheme', 'Pay Scheme'),
        'level_id': fields.many2one('ng.state.payroll.level', 'Grade'),
        'department_id': fields.many2one('hr.department', 'MDA', required=True),
    }
    
    _defaults = {
        'gross_income': 0,
        'taxable_income': 0,
        'net_income': 0,
        'leave_allowance': 0,
        'paye_tax': 0,
        'paye_tax_annual': 0,
        'active': True,
        'resolve': False,
        'retiring': False,
    }                                          

    def list_payroll_item_lines(self, cr, uid, item_id, context=None):
        _logger.info("Calling list_payroll_item_lines")
        _logger.info("User ID=%s, Item ID=%s", uid, item_id)

        item_lines = []
        if uid and item_id:
            cr.execute("select name,amount from ng_state_payroll_payroll_item_line where item_id=" + str(item_id))
            item_lines = cr.fetchall()
        else:
            _logger.info("No matching employee found for ID %d", uid)
        
        _logger.info("Lines=%s", item_lines)   
        return item_lines
   
    @api.depends('payroll_id', 'employee_id')   
    def _compute_payment_items(self):
        for payroll_item in self:
            payroll_item.payment_item_ids = self.env['ng.state.payroll.scenario.payment'].search([('employee_id.id', '=', payroll_item.employee_id.id), ('scenario_id.payroll_id.id', '=', payroll_item.payroll_id.id)])        
                        
class ng_state_payroll_payroll_item_line(models.Model):
    '''
    Payroll Item Line
    '''
    _name = "ng.state.payroll.payroll.item.line"
    _description = 'Payroll Item Line'

    _columns = {
        'code': fields.char('Code', help='Line Code', required=False),
        'name': fields.char('Name', help='Line Name', required=False),
        'amount': fields.float('Amount', help='Amount', required=True),
        'item_id': fields.many2one('ng.state.payroll.payroll.item', 'Payroll Item'),
    }
                        
class ng_state_payroll_pension_item_line(models.Model):
    '''
    Pension Item Line
    '''
    _name = "ng.state.payroll.pension.item.line"
    _description = 'Pension Item Line'

    _columns = {
        'code': fields.char('Code', help='Line Code', required=False),
        'name': fields.char('Name', help='Line Name', required=False),
        'amount': fields.float('Amount', help='Amount', required=True),
        'item_id': fields.many2one('ng.state.payroll.pension.item', 'Pension Item'),
    }
    
class ng_state_payroll_payroll_signoff(models.Model):
    '''
    Payroll Sign-Off
    '''
    _name = "ng.state.payroll.payroll.signoff"
    _description = 'Payroll Sign-Off'

    _columns = {
        'payroll_id': fields.many2one('ng.state.payroll.payroll', 'Payroll', required=True),
        'user_id': fields.many2one('res.users', 'Payroll Officer', required=True, domain="[('groups_id.name','=','Payroll Officer')]"),
        'signed_off': fields.boolean('Closed', help='Sign-Off closed status', required=True),
        'pos_order': fields.integer('Order', help='Order', required=True),
    }
    
    _defaults = {
        'signed_off': False,
    }                                          

    def name_get(self, cr, uid, ids, context=None):
        if not ids:
            return []
        res = []
        for r in self.read(cr, uid, ids, ['id', 'payroll_id', 'user_id'], context):
            aux = " ("
            if r['payroll_id']:
                aux += r['payroll_id'][1]
                # why translate a date? I think is a mistake, the _() function must have a 
                # known string, example _("the start date is %s") % r['from_date']
    
            aux +=  ' - '
            if r['user_id']:
                aux += r['user_id'][1] # same question
    
            aux += ')'
    
            # aux is the name items for the r['id']
            res.append((r['id'], aux))  # append add the tuple (r['id'], aux) in the list res
    
        return res
        #Open create form with current month date range
    def name_search(self, cr, user, name='', args=None, operator='ilike', context=None, limit=100):
        if not args:
            args = []
        if name:
            ids = self.search(cr, user, [('payroll_id.name','=',name)]+ args, limit=limit, context=context)
            if not ids:
                # Do not merge the 2 next lines into one single search, SQL search performance would be abysmal
                # on a database with thousands of matching products, due to the huge merge+unique needed for the
                # OR operator (and given the fact that the 'name' lookup results come from the ir.translation table
                # Performing a quick memory merge of ids in Python will give much better performance
                ids = set()
                ids.update(self.search(cr, user, args + [('user_id',operator,name)], limit=limit, context=context))
                ids = list(ids)
            if not ids:
                ptrn = re.compile('(\[(.*?)\])')
                res = ptrn.search(name)
                if res:
                    ids = self.search(cr, user, [('id','=', res.group(2))] + args, limit=limit, context=context)
        else:
            ids = self.search(cr, user, args, limit=limit, context=context)
        result = self.name_get(cr, user, ids, context=context)
        return result
   
class ng_state_payroll_template_earning(models.Model):
    '''
    Payroll Employee Earning/Deduction Upload Earning Template
    '''
    _name = "ng.state.payroll.template.earning"
    _description = 'Payroll Employee Earning/Deduction Upload Earning Template'            
    

    _columns = {
        'name': fields.char('Name', help='Deduction Name', required=True),
        'permanent': fields.boolean('Permanent', help='Permanent'),
    }
    
    _defaults = {
        'permanent': False,
    }     
   
class ng_state_payroll_template_deduction(models.Model):
    '''
    Payroll Employee Earning/Deduction Upload Deduction Template
    '''
    _name = "ng.state.payroll.template.deduction"
    _description = 'Payroll Employee Earning/Deduction Upload Deduction Template'            
    

    _columns = {
        'name': fields.char('Name', help='Deduction Name', required=True),
        'permanent': fields.boolean('Permanent', help='Permanent'),
        'relief': fields.boolean('Relief', help='Forming part of CRF Relief'),
        'income_deduction': fields.boolean('Income Deduction', help='Deducted from Income before CRA 20% calculation'),
    }
    
    _defaults = {
        'permanent': False,
        'relief': False,
        'income_deduction': False,
    }     

class ng_state_payroll_earnded_upload(models.Model):
    '''
    Payroll Employee Earning/Deduction Upload
    '''
    _name = "ng.state.payroll.earnded.upload"
    _description = 'Payroll Employee Earning/Deduction Upload'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'name': fields.char('Upload Name', help='Upload Name', required=True),
        'upload_file': fields.binary('Nonstandard Earnings/Deductions File'),
        'deduction': fields.boolean('Deduction Upload', help='Records are deductions when true or earnings when false'),
        'arrears': fields.boolean('Arrears', help='Records are arrears'),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'user_id': fields.many2one('res.users', 'Payroll Officer', readonly=True, required=True, domain="[('groups_id.name','=','Payroll Officer')]"),
        'calendar_id': fields.many2one('ng.state.payroll.calendar', 'Calendar', track_visibility='onchange', required=False),
        'notify_emails': fields.char('Notify Email', help='Comma separated email recipients for event notification', required=True),
        'nonstd_earnings': fields.many2many('ng.state.payroll.earning.nonstd', 'rel_upload_std_earning', 'upload_id','earning_id', 'Nonstandard Earnings'), 
        'nonstd_deductions': fields.many2many('ng.state.payroll.deduction.nonstd', 'rel_upload_std_deduction', 'upload_id','deduction_id', 'Nonstandard Deductions'), 
    }

    _rec_name = 'date'
    
    _defaults = {
        'state': 'draft',
        'date': date.today(),
        'user_id': lambda s, cr, uid, c: uid,
    }
       
    _track = {
        'state': {
            'ng_state_payroll_earnded_upload.mt_alert_earnded_upload_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_earnded_upload.mt_alert_earnded_upload_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }
    
    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        _logger.info("_needaction_domain_get - %s", users_obj)

        if users_obj.has_group(cr, uid, 'hr_payroll_nigerian_state.group_payroll_administrator'):
            _logger.info("_needaction_domain_get - is Payroll Administrator")
            domain = [('state', '=', 'confirm')]
            return domain

        return False

    
    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Earning/Deduction Upload action!'),
                    _('Earning/Deduction Upload action has been initiated. Either cancel the earnded_upload action or create another to undo it.')
                )

        return super(ng_state_payroll_earnded_upload, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for o in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                o.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def earnded_upload_state_confirm(self, cr, uid, ids, context=None):
        #TODO Process file, select distinct by name and create templates for earnings/deductions
        _logger.info("before state_confirm - %d", uid)
        self.write(cr, uid, ids, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True
           
    @api.multi
    def revert(self):
        _logger.info("Calling revert..id = %s", self.id)
        #if self.env.user.has_group('hr_payroll_nigerian_state.group_payroll_admin'):

        self.env.cr.execute("update ng_state_payroll_earnded_upload set state='draft' where id=" + str(self.id))
        if not self.deduction:
            self.env.cr.execute("delete from ng_state_payroll_earning_nonstd where id in (select earning_id from rel_upload_std_earning where upload_id=" + str(self.id) + ")")
            self.env.cr.execute("delete from rel_upload_std_earning where upload_id=" + str(self.id))
        else:
            self.env.cr.execute("delete from ng_state_payroll_deduction_nonstd where id in (select deduction_id from rel_upload_std_deduction where upload_id=" + str(self.id) + ")")
            self.env.cr.execute("delete from rel_upload_std_deduction where upload_id=" + str(self.id))
        self.env.invalidate_all()

    def try_confirmed_earnded_upload_actions(self, cr, uid, context=None):
        _logger.info("Running try_confirmed_earnded_upload_actions cron-job...")
        employee_obj = self.pool.get('hr.employee')
        user_obj = self.pool.get('res.users')
        upload_obj = self.pool.get('ng.state.payroll.earnded.upload')
        today = datetime.now().date()

        cr.execute('deallocate all')
        cr.execute('prepare insert_nonstd_earning (int,text,numeric,bool,timestamp,timestamp,bool,bool,int,int) as insert into ng_state_payroll_earning_nonstd (employee_id,name,amount,active,create_date,write_date,permanent,taxable,create_uid,write_uid) values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10) returning id')            
        cr.execute('prepare insert_nonstd_deduction (int,text,numeric,bool,timestamp,timestamp,bool,bool,bool,int,int) as insert into ng_state_payroll_deduction_nonstd (employee_id,name,amount,active,create_date,write_date,permanent,relief,income_deduction,create_uid,write_uid) values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) returning id')            
        cr.execute('prepare insert_earning_nonstd_calendar (int, int) as insert into rel_earning_nonstd_calendar (calendar_id,earning_id) values ($1,$2)')
        cr.execute('prepare insert_deduction_nonstd_calendar (int, int) as insert into rel_deduction_nonstd_calendar (calendar_id,deduction_id) values ($1,$2)')            
        cr.execute('prepare insert_upload_std_earning (int, int) as insert into rel_upload_std_earning (upload_id,earning_id) values ($1,$2)')
        cr.execute('prepare insert_upload_std_deduction (int, int) as insert into rel_upload_std_deduction (upload_id,deduction_id) values ($1,$2)')

        upload_ids = upload_obj.search(cr, uid, [('state', '=', 'confirm')], context=context)
        
        for upload in self.browse(cr, uid, upload_ids, context=context):
            if upload.upload_file != None and datetime.strptime(upload.date, DEFAULT_SERVER_DATE_FORMAT).date() <= today and upload.state == 'confirm':
                exception_list = []
                payroll_officer = user_obj.browse(cr, uid, upload.user_id.id, context=context)
                data_file = base64.decodestring(upload.upload_file)
                permanent = 'f'
                if not upload.calendar_id:
                    permanent = 't'
                wb = open_workbook(file_contents=data_file)
                for s in wb.sheets():
                    _logger.info("Number of sheets: %d", len(wb.sheets()))
                    _logger.info("Number of records: %d", s.nrows)
                    for row in range(s.nrows):
                        exec_inserts = False
                        if row > 0: #Skip first row
                            data_row = []
                            for col in range(s.ncols):
                                value = (s.cell(row, col).value)
                                data_row.append(value)
                            cr.execute("select id from hr_employee where employee_no='" + str(data_row[0]).strip().replace("'", "") + "'")
                            employee_id = cr.fetchall()                            
                            item_id = False
                            employee = False
                            if employee_id:
                                employee = employee_obj.browse(cr, uid, employee_id[0], context=context)
                            if payroll_officer.domain_mdas:
                                if employee and employee.department_id.id in payroll_officer.domain_mdas.ids:
                                    exec_inserts = True
                                else:
                                    if employee:
                                        exception_list.append({'employee_no':data_row[0],'description':data_row[1],'amount':data_row[2],'error':'Wrong MDA - ' + str(employee.department_id.name)})
                                    else:
                                        if len(data_row) == 3:
                                            exception_list.append({'employee_no':data_row[0],'description':data_row[1],'amount':data_row[2],'error':'No employee found.'})
                                        else:
                                            exception_list.append({'employee_no':data_row[0],'description':'','amount':'','error':'Wrong number of spreadsheet columns'})
                            else:
                                exec_inserts = True
                            if exec_inserts:
                                if employee_id and len(employee_id) == 1 and len(data_row) == 3:
                                    description = data_row[1]
                                    amount = str(data_row[2]).strip().replace(',','')
                                    if upload.arrears:
                                        description = 'ARREARS - ' + description
                                    if upload.deduction:
                                        relief = 'f'
                                        if str(data_row[1]).upper().startswith('NHF') or str(data_row[1]).upper().startswith('PENSION'):
                                            relief = 't'
                                        if is_number(amount):
                                            cr.execute('execute insert_nonstd_deduction(%s,%s,%s,%s,current_timestamp,current_timestamp,%s,%s,%s,%s,%s)', (employee[0].id,description,amount,'t',permanent,relief,'f',uid,uid))
                                            item_id = cr.fetchone()
                                            cr.execute('execute insert_upload_std_deduction(%s,%s)', (upload.id,item_id))
                                            if upload.calendar_id:
                                                cr.execute('execute insert_deduction_nonstd_calendar(%s,%s)', (upload.calendar_id.id,item_id))
                                        else:
                                            exception_list.append({'employee_no':data_row[0],'description':'','amount':'','error':'Amount not valid: ' + amount})
                                    else:
                                        if is_number(amount):
                                            cr.execute('execute insert_nonstd_earning(%s,%s,%s,%s,current_timestamp,current_timestamp,%s,%s,%s,%s)', (employee[0].id,description,amount,'t',permanent,'t',uid,uid))
                                            item_id = cr.fetchone()
                                            cr.execute('execute insert_upload_std_earning(%s,%s)', (upload.id,item_id))
                                            if upload.calendar_id:
                                                cr.execute('execute insert_earning_nonstd_calendar(%s,%s)', (upload.calendar_id.id,item_id))
                                        else:
                                            exception_list.append({'employee_no':data_row[0],'description':'','amount':'','error':'Amount not valid: ' + amount})
                                else:
                                    _logger.info("%s - Employee Length: %d, Data Length: %d", data_row[0],len(employee_id),len(data_row))
                                    if len(employee_id) != 1:
                                        exception_list.append({'employee_no':data_row[0],'description':'','amount':'','error':'No employee found for Employee Number'}) 
                                    else:
                                        exception_list.append({'employee_no':data_row[0],'description':'','amount':'','error':'Wrong number of spreadsheet columns'})
                cr.execute("update ng_state_payroll_earnded_upload set state='approved' where id=" + str(upload.id))                
                if len(exception_list) > 0:
                    with open('/tmp/' + 'nonstd_earnded_upload_exceptions_' + str(upload.id) + '.csv', 'w') as csvfile:
                        fieldnames = ['employee_no', 'description', 'amount','error']
                        if not upload.deduction:
                            fieldnames = ['employee_no', 'description', 'amount','permanent','relief','income_deduction','error']
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        writer.writeheader()
                        writer.writerows(exception_list)
                        csvfile.close()

                upload_type = 'earnings'
                if upload.deduction:
                    upload_type = 'deductions'
                message = "Dear Sir/Madam,\nUpload of nonstandard " + upload_type + " file has completed.\n\nThank you.\n"
                message_exception = "\nPS: There were " + str(len(exception_list)) + " exceptions."
                smtp_obj = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
                smtp_obj.ehlo()
                #smtp_obj.starttls()
                #smtp_obj.ehlo()
                smtp_obj.login(user="osun.payroll@gmail.com", password="p@55w0rd1939")
                sender = 'osun.payroll@gmail.com'
                receivers = upload.notify_emails #Comma separated email addresses
                msg = MIMEMultipart()
                msg['Subject'] = 'Upload Completed' 
                msg['From'] = sender
                #msg['To'] = ', '.join(receivers)
                msg['To'] = receivers
                                 
                part = False
                if len(exception_list) > 0:
                    part = MIMEBase('application', "octet-stream")
                    part.set_payload(open('/tmp/' + 'nonstd_earnded_upload_exceptions_' + str(upload.id) + '.csv', "rb").read())
                    Encoders.encode_base64(part)                            
                    part.add_header('Content-Disposition', 'attachment; filename="nonstd_earnded_upload_exceptions_' + str(upload.id) + '.csv"')
                    message = message + message_exception
                    msg.attach(MIMEText(message))
                                
                    if part:
                        msg.attach(part)
                                                        
                smtp_obj.sendmail(sender, receivers, msg.as_string())         
                _logger.info("Email successfully sent to: " + receivers)   
            else:
                return False
        cr.commit()

        return True

class ng_state_payroll_certificate_upload(models.Model):
    '''
    Employee Certificate Upload
    '''
    _name = "ng.state.payroll.certificate.upload"
    _description = 'Employee Certificate Upload'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'name': fields.char('Upload Name', help='Upload Name', required=True),
        'upload_file': fields.binary('Certificate Upload File'),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'user_id': fields.many2one('res.users', 'HR Manager', readonly=True, required=True, domain="[('groups_id.name','=','Manager')]"),
        'notify_emails': fields.char('Notify Email', help='Comma separated email recipients for event notification', required=True),
        'cert_uploads': fields.many2many('ng.state.payroll.certification', 'rel_upload_employee_cert', 'upload_id','certification_id', 'Uploaded Certificates'), 
    }
 
    _rec_name = 'date'
    
    _defaults = {
        'state': 'draft',
        'date': date.today(),
        'user_id': lambda s, cr, uid, c: uid,
    }
       
    _track = {
        'state': {
            'ng_state_payroll_certificate_upload.mt_alert_certificate_upload_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_certificate_upload.mt_alert_certificate_upload_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }
    
    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        _logger.info("_needaction_domain_get - %s", users_obj)

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            _logger.info("_needaction_domain_get - is HR Manager")
            domain = [('state', '=', 'confirm')]
            return domain

        return False
    
    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Earning/Deduction Upload action!'),
                    _('Earning/Deduction Upload action has been initiated. Either cancel the certificate_upload action or create another to undo it.')
                )

        return super(ng_state_payroll_certificate_upload, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for o in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                o.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def certificate_upload_state_confirm(self, cr, uid, ids, context=None):
        #TODO Process file, select distinct by name and create templates for earnings/deductions
        _logger.info("before state_confirm - %d", uid)
        self.write(cr, uid, ids, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def try_confirmed_certificate_upload_actions(self, cr, uid, context=None):
        _logger.info("Running try_confirmed_certificate_upload_actions cron-job...")
        employee_obj = self.pool.get('hr.employee')
#         user_obj = self.pool.get('res.users')
        upload_obj = self.pool.get('ng.state.payroll.certificate.upload')
        today = datetime.now().date()

        cr.execute('deallocate all')
        cr.execute("prepare insert_certification (int,int,int,int,date,bool,text,timestamp,timestamp,int,int) as insert into ng_state_payroll_certification (employee_id,upload_id,certificate_id,course_id,date,active,state,create_date,write_date,create_uid,write_uid) values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) returning id")            
        cr.execute("prepare insert_certification2 (int,int,int,date,bool,text,timestamp,timestamp,int,int) as insert into ng_state_payroll_certification (employee_id,upload_id,certificate_id,date,active,state,create_date,write_date,create_uid,write_uid) values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10) returning id")            
        cr.execute("prepare insert_certification3 (int,int,int,int,bool,text,timestamp,timestamp,int,int) as insert into ng_state_payroll_certification (employee_id,upload_id,certificate_id,course_id,active,state,create_date,write_date,create_uid,write_uid) values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10) returning id")            
        cr.execute("prepare insert_certification4 (int,int,int,bool,text,timestamp,timestamp,int,int) as insert into ng_state_payroll_certification (employee_id,upload_id,certificate_id,active,state,create_date,write_date,create_uid,write_uid) values ($1,$2,$3,$4,$5,$6,$7,$8,$9) returning id")            
        cr.execute('prepare insert_upload_employee_cert (int, int) as insert into rel_upload_employee_cert (upload_id,certification_id) values ($1,$2)')
        cr.execute('prepare insert_employee_certification (int, int) as insert into rel_employee_certification (employee_id,certification_id) values ($1,$2)')

        upload_ids = upload_obj.search(cr, uid, [('state', '=', 'confirm')], context=context)
        
        for upload in self.browse(cr, uid, upload_ids, context=context):
            if upload.upload_file != None and datetime.strptime(upload.date, DEFAULT_SERVER_DATE_FORMAT).date() <= today and upload.state == 'confirm':
                exception_list = []
#                 hr_officer = user_obj.browse(cr, uid, upload.user_id.id, context=context)
                data_file = base64.decodestring(upload.upload_file)
                wb = open_workbook(file_contents=data_file)
                warnings = 0
                for s in wb.sheets():
                    _logger.info("Number of sheets: %d", len(wb.sheets()))
                    _logger.info("Number of records: %d", s.nrows)
                    for row in range(s.nrows):
                        if row > 0: #Skip first row
                            data_row = []
                            for col in range(s.ncols):
                                value = (s.cell(row, col).value)
                                data_row.append(value)
                            cr.execute("select id from hr_employee where employee_no='" + str(data_row[0]).strip().replace("'", "") + "'")
                            employee_id = cr.fetchall()                            
                            certification_id = False
                            employee = False
                            if employee_id:
                                employee = employee_obj.browse(cr, uid, employee_id[0], context=context)
                            if employee_id and len(employee_id) == 1 and (len(data_row) == 3 or len(data_row) == 4):
                                certificate_name = str(data_row[1]).strip().replace("'", "")
                                course_name = str(data_row[2]).strip().replace("'", "")

                                cr.execute("select id from ng_state_payroll_certificate where name='" + certificate_name + "'")
                                certificate_id = cr.fetchall()                            

                                cr.execute("select id from ng_state_payroll_certcourse where name='" + course_name + "'")
                                course_id = cr.fetchall()                            

                                cert_date = False
                                if len(data_row) == 4:
                                    cert_date = str(data_row[3]).strip().replace(',','')
                                    try:
                                        datetime.strptime(cert_date, '%Y-%m-%d')
                                    except ValueError:
                                        _logger.error("Incorrect date format: '" + cert_date + "'; should be YYYY-MM-DD")
                                        exception_list.append({'employee_no':data_row[0],'certificate_name':certificate_name,'course_name':course_name,'date':cert_date,'warning':'Incorrect date format: \'' + cert_date + '\'; should be YYYY-MM-DD'})
                                        warnings += 1                                    
                                        cert_date = False
                                
                                if not course_id:
                                    course_id = False
                                else:
                                    course_id = course_id[0]        
                                if certificate_id:
                                    if cert_date:
                                        if course_id:
                                            cr.execute('execute insert_certification(%s,%s,%s,%s,%s,%s,%s,current_timestamp,current_timestamp,%s,%s)', (employee[0].id,upload.id,certificate_id[0],course_id,cert_date,'t','unconfirmed',uid,uid))
                                        else:
                                            cr.execute('execute insert_certification2(%s,%s,%s,%s,%s,%s,current_timestamp,current_timestamp,%s,%s)', (employee[0].id,upload.id,certificate_id[0],cert_date,'t','unconfirmed',uid,uid))
                                    else:
                                        if course_id:
                                            cr.execute('execute insert_certification3(%s,%s,%s,%s,%s,%s,current_timestamp,current_timestamp,%s,%s)', (employee[0].id,upload.id,certificate_id[0],course_id,'t','unconfirmed',uid,uid))
                                        else:
                                            cr.execute('execute insert_certification4(%s,%s,%s,%s,%s,current_timestamp,current_timestamp,%s,%s)', (employee[0].id,upload.id,certificate_id[0],'t','unconfirmed',uid,uid))
                                    certification_id = cr.fetchone()
                                    cr.execute('execute insert_upload_employee_cert(%s,%s)', (upload.id,certification_id))
                                    cr.execute('execute insert_employee_certification(%s,%s)', (employee_id[0],certification_id))
                                else:
                                    exception_list.append({'employee_no':data_row[0],'certificate_name':certificate_name,'course_name':course_name,'date':cert_date,'error':'No certificate found for Certificate Name ' + certificate_name})
                            else:
                                _logger.info("%s - Employee Length: %d, Data Length: %d", data_row[0],len(employee_id),len(data_row))
                                if len(employee_id) != 1:
                                    exception_list.append({'employee_no':data_row[0],'certificate_name':'','course_name':'','date':'','error':'No employee found for Employee Number'}) 
                                else:
                                    exception_list.append({'employee_no':data_row[0],'certificate_name':'','course_name':'','date':'','error':'Wrong number of spreadsheet columns'})
                cr.execute("update ng_state_payroll_certificate_upload set state='approved' where id=" + str(upload.id))                
                if len(exception_list) > 0:
                    with open('/tmp/' + 'certificate_upload_exceptions_' + str(upload.id) + '.csv', 'w') as csvfile:
                        fieldnames = ['employee_no', 'certificate_name', 'course_name','date','error','warning']
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        writer.writeheader()
                        writer.writerows(exception_list)
                        csvfile.close()

                message = "Dear Sir/Madam,\nUpload of certificate file has completed.\n\nThank you.\n"
                message_exception = "\nPS: There were " + str(len(exception_list)) + " exceptions and " + str(warnings) + " warnings."
                smtp_obj = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
                smtp_obj.ehlo()
                #smtp_obj.starttls()
                #smtp_obj.ehlo()
                smtp_obj.login(user="osun.payroll@gmail.com", password="p@55w0rd1939")
                sender = 'osun.payroll@gmail.com'
                receivers = upload.notify_emails #Comma separated email addresses
                msg = MIMEMultipart()
                msg['Subject'] = 'Upload Completed' 
                msg['From'] = sender
                #msg['To'] = ', '.join(receivers)
                msg['To'] = receivers
                                 
                part = False
                if len(exception_list) > 0:
                    part = MIMEBase('application', "octet-stream")
                    part.set_payload(open('/tmp/' + 'certificate_upload_exceptions_' + str(upload.id) + '.csv', "rb").read())
                    Encoders.encode_base64(part)                            
                    part.add_header('Content-Disposition', 'attachment; filename="nonstd_certificate_upload_exceptions_' + str(upload.id) + '.csv"')
                    message = message + message_exception
                    msg.attach(MIMEText(message))
                                
                    if part:
                        msg.attach(part)
                                                        
                smtp_obj.sendmail(sender, receivers, msg.as_string())         
                _logger.info("Email successfully sent to: " + receivers)   
            else:
                return False
        cr.commit()

        return True

class ng_state_payroll_jobdesc_upload(models.Model):
    '''
    Employee Job Description Upload
    '''
    _name = "ng.state.payroll.jobdesc.upload"
    _description = 'Employee Job Description Upload'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'name': fields.char('Upload Name', help='Upload Name', required=True),
        'upload_file': fields.binary('Job Description Upload File'),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'user_id': fields.many2one('res.users', 'HR Manager', readonly=True, required=True, domain="[('groups_id.name','=','Manager')]"),
        'notify_emails': fields.char('Notify Email', help='Comma separated email recipients for event notification', required=True),
        'jobdesc_uploads': fields.many2many('hr.employee', 'rel_upload_employee_jobdesc', 'upload_id','employee_id', 'Uploaded Job Descriptions'), 
    }
 
    _rec_name = 'date'
    
    _defaults = {
        'state': 'draft',
        'date': date.today(),
        'user_id': lambda s, cr, uid, c: uid,
    }
       
    _track = {
        'state': {
            'ng_state_payroll_jobdesc_upload.mt_alert_jobdesc_upload_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_jobdesc_upload.mt_alert_jobdesc_upload_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }
    
    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        _logger.info("_needaction_domain_get - %s", users_obj)

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            _logger.info("_needaction_domain_get - is HR Manager")
            domain = [('state', '=', 'confirm')]
            return domain

        return False

    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Earning/Deduction Upload action!'),
                    _('Earning/Deduction Upload action has been initiated. Either cancel the jobdesc_upload action or create another to undo it.')
                )

        return super(ng_state_payroll_jobdesc_upload, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for o in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                o.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def jobdesc_upload_state_confirm(self, cr, uid, ids, context=None):
        _logger.info("before state_confirm - %d", uid)
        self.write(cr, uid, ids, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def try_confirmed_jobdesc_upload_actions(self, cr, uid, context=None):
        _logger.info("Running try_confirmed_jobdesc_upload_actions cron-job...")
        employee_obj = self.pool.get('hr.employee')
        upload_obj = self.pool.get('ng.state.payroll.jobdesc.upload')
        today = datetime.now().date()

        cr.execute('deallocate all')
        cr.execute('prepare update_jobdesc (text, int) as update hr_employee set job_description = $1 where id = $2')

        upload_ids = upload_obj.search(cr, uid, [('state', '=', 'confirm')], context=context)
        
        for upload in self.browse(cr, uid, upload_ids, context=context):
            if upload.upload_file != None and datetime.strptime(upload.date, DEFAULT_SERVER_DATE_FORMAT).date() <= today and upload.state == 'confirm':
                exception_list = []
                data_file = base64.decodestring(upload.upload_file)
                wb = open_workbook(file_contents=data_file)
                for s in wb.sheets():
                    _logger.info("Number of sheets: %d", len(wb.sheets()))
                    _logger.info("Number of records: %d", s.nrows)
                    for row in range(s.nrows):
                        if row > 0: #Skip first row
                            data_row = []
                            for col in range(s.ncols):
                                value = (s.cell(row, col).value)
                                data_row.append(value)
                            cr.execute("select id from hr_employee where employee_no='" + str(data_row[0]).strip().replace("'", "") + "'")
                            employee_id = cr.fetchall()                            
                            employee = False
                            if employee_id:
                                employee = employee_obj.browse(cr, uid, employee_id[0], context=context)
                            if employee_id and len(employee_id) == 1 and len(data_row) == 2:
                                job_desc = unicode(data_row[1]).strip().replace("'", "''")
                                cr.execute('execute update_jobdesc(%s,%s)', (job_desc,employee_id[0]))
                            else:
                                _logger.info("%s - Employee Length: %d, Data Length: %d", data_row[0],len(employee_id),len(data_row))
                                if len(employee_id) != 1:
                                    exception_list.append({'employee_no':data_row[0],'job_desc':'','error':'No employee found for Employee Number'}) 
                                else:
                                    exception_list.append({'employee_no':data_row[0],'job_desc':'','error':'Wrong number of spreadsheet columns'})
                cr.execute("update ng_state_payroll_jobdesc_upload set state='approved' where id=" + str(upload.id))                
                if len(exception_list) > 0:
                    with open('/tmp/' + 'jobdesc_upload_exceptions_' + str(upload.id) + '.csv', 'w') as csvfile:
                        fieldnames = ['employee_no', 'job_desc','error']
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        writer.writeheader()
                        writer.writerows(exception_list)
                        csvfile.close()

                message = "Dear Sir/Madam,\nUpload of jobdesc file has completed.\n\nThank you.\n"
                message_exception = "\nPS: There were " + str(len(exception_list)) + " exceptions."
                smtp_obj = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
                smtp_obj.ehlo()
                #smtp_obj.starttls()
                #smtp_obj.ehlo()
                smtp_obj.login(user="osun.payroll@gmail.com", password="p@55w0rd1939")
                sender = 'osun.payroll@gmail.com'
                receivers = upload.notify_emails #Comma separated email addresses
                msg = MIMEMultipart()
                msg['Subject'] = 'Upload Completed' 
                msg['From'] = sender
                #msg['To'] = ', '.join(receivers)
                msg['To'] = receivers
                                 
                part = False
                if len(exception_list) > 0:
                    part = MIMEBase('application', "octet-stream")
                    part.set_payload(open('/tmp/' + 'jobdesc_upload_exceptions_' + str(upload.id) + '.csv', "rb").read())
                    Encoders.encode_base64(part)                            
                    part.add_header('Content-Disposition', 'attachment; filename="jobdesc_upload_exceptions_' + str(upload.id) + '.csv"')
                    message = message + message_exception
                    msg.attach(MIMEText(message))
                                
                    if part:
                        msg.attach(part)
                                                        
                smtp_obj.sendmail(sender, receivers, msg.as_string())         
                _logger.info("Email successfully sent to: " + receivers)   
            else:
                return False
        cr.commit()

        return True

class ng_state_payroll_bvn_upload(models.Model):
    '''
    Employee BVN Upload
    '''
    _name = "ng.state.payroll.bvn.upload"
    _description = 'Employee BVN Upload'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'name': fields.char('Upload Name', help='Upload Name', required=True),
        'upload_file': fields.binary('BVN Upload File'),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'user_id': fields.many2one('res.users', 'HR Manager', readonly=True, required=True, domain="[('groups_id.name','=','Manager')]"),
        'notify_emails': fields.char('Notify Email', help='Comma separated email recipients for event notification', required=True),
        'bvn_uploads': fields.many2many('hr.employee', 'rel_upload_employee_bvn', 'upload_id','employee_id', 'Uploaded BVNs'), 
    }
 
    _rec_name = 'date'
    
    _defaults = {
        'state': 'draft',
        'date': date.today(),
        'user_id': lambda s, cr, uid, c: uid,
    }
       
    _track = {
        'state': {
            'ng_state_payroll_bvn_upload.mt_alert_bvn_upload_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_bvn_upload.mt_alert_bvn_upload_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }
    
    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        _logger.info("_needaction_domain_get - %s", users_obj)

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            _logger.info("_needaction_domain_get - is HR Manager")
            domain = [('state', '=', 'confirm')]
            return domain

        return False

    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Earning/Deduction Upload action!'),
                    _('Earning/Deduction Upload action has been initiated. Either cancel the bvn_upload action or create another to undo it.')
                )

        return super(ng_state_payroll_bvn_upload, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for o in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                o.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def bvn_upload_state_confirm(self, cr, uid, ids, context=None):
        _logger.info("before state_confirm - %d", uid)
        self.write(cr, uid, ids, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def try_confirmed_bvn_upload_actions(self, cr, uid, context=None):
        _logger.info("Running try_confirmed_bvn_upload_actions cron-job...")
        employee_obj = self.pool.get('hr.employee')
        upload_obj = self.pool.get('ng.state.payroll.bvn.upload')
        today = datetime.now().date()

        cr.execute('deallocate all')
        cr.execute('prepare update_bvn (text, int) as update hr_employee set bvn = $1 where id = $2')

        upload_ids = upload_obj.search(cr, uid, [('state', '=', 'confirm')], context=context)
        
        for upload in self.browse(cr, uid, upload_ids, context=context):
            if upload.upload_file != None and datetime.strptime(upload.date, DEFAULT_SERVER_DATE_FORMAT).date() <= today and upload.state == 'confirm':
                exception_list = []
                data_file = base64.decodestring(upload.upload_file)
                wb = open_workbook(file_contents=data_file)
                for s in wb.sheets():
                    _logger.info("Number of sheets: %d", len(wb.sheets()))
                    _logger.info("Number of records: %d", s.nrows)
                    for row in range(s.nrows):
                        if row > 0: #Skip first row
                            data_row = []
                            for col in range(s.ncols):
                                value = (s.cell(row, col).value)
                                data_row.append(value)
                            cr.execute("select id from hr_employee where employee_no='" + str(data_row[0]).strip().replace("'", "") + "'")
                            employee_id = cr.fetchall()                            
                            employee = False
                            if employee_id:
                                employee = employee_obj.browse(cr, uid, employee_id[0], context=context)
                            if employee_id and len(employee_id) == 1 and len(data_row) == 2:
                                bvn = str(data_row[1]).strip().replace("'", "").replace(".0", "")
                                cr.execute("select id from hr_employee where bvn='" + bvn + "'")
                                employee_id_check = cr.fetchall()
                                if not employee_id_check:
                                    cr.execute('execute update_bvn(%s,%s)', (bvn,employee_id[0]))
                                else:
                                    exception_list.append({'employee_no':data_row[0],'bvn':'','error':'BVN already exists; must be unique - ' + bvn}) 
                            else:
                                _logger.info("%s - Employee Length: %d, Data Length: %d", data_row[0],len(employee_id),len(data_row))
                                if len(employee_id) != 1:
                                    exception_list.append({'employee_no':data_row[0],'bvn':'','error':'No employee found for Employee Number'}) 
                                else:
                                    exception_list.append({'employee_no':data_row[0],'bvn':'','error':'Wrong number of spreadsheet columns'})
                cr.execute("update ng_state_payroll_bvn_upload set state='approved' where id=" + str(upload.id))                
                if len(exception_list) > 0:
                    with open('/tmp/' + 'bvn_upload_exceptions_' + str(upload.id) + '.csv', 'w') as csvfile:
                        fieldnames = ['employee_no', 'bvn','error']
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        writer.writeheader()
                        writer.writerows(exception_list)
                        csvfile.close()

                message = "Dear Sir/Madam,\nUpload of BVN file has completed.\n\nThank you.\n"
                message_exception = "\nPS: There were " + str(len(exception_list)) + " exceptions."
                smtp_obj = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
                smtp_obj.ehlo()
                #smtp_obj.starttls()
                #smtp_obj.ehlo()
                smtp_obj.login(user="osun.payroll@gmail.com", password="p@55w0rd1939")
                sender = 'osun.payroll@gmail.com'
                receivers = upload.notify_emails #Comma separated email addresses
                msg = MIMEMultipart()
                msg['Subject'] = 'Upload Completed' 
                msg['From'] = sender
                #msg['To'] = ', '.join(receivers)
                msg['To'] = receivers
                                 
                part = False
                if len(exception_list) > 0:
                    part = MIMEBase('application', "octet-stream")
                    part.set_payload(open('/tmp/' + 'bvn_upload_exceptions_' + str(upload.id) + '.csv', "rb").read())
                    Encoders.encode_base64(part)                            
                    part.add_header('Content-Disposition', 'attachment; filename="bvn_upload_exceptions_' + str(upload.id) + '.csv"')
                    message = message + message_exception
                    msg.attach(MIMEText(message))
                                
                    if part:
                        msg.attach(part)
                                                        
                smtp_obj.sendmail(sender, receivers, msg.as_string())         
                _logger.info("Email successfully sent to: " + receivers)   
            else:
                return False
        cr.commit()

        return True

class ng_state_payroll_designation_upload(models.Model):
    '''
    Employee Designation Upload
    '''
    _name = "ng.state.payroll.designation.upload"
    _description = 'Employee Designation Upload'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'name': fields.char('Upload Name', help='Upload Name', required=True),
        'upload_file': fields.binary('Designation Upload File'),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'user_id': fields.many2one('res.users', 'HR Manager', readonly=True, required=True, domain="[('groups_id.name','=','Manager')]"),
        'notify_emails': fields.char('Notify Email', help='Comma separated email recipients for event notification', required=True),
        'designation_uploads': fields.many2many('hr.employee', 'rel_upload_employee_designation', 'upload_id','employee_id', 'Uploaded Designations'), 
    }
 
    _rec_name = 'date'
    
    _defaults = {
        'state': 'draft',
        'date': date.today(),
        'user_id': lambda s, cr, uid, c: uid,
    }
       
    _track = {
        'state': {
            'ng_state_payroll_designation_upload.mt_alert_designation_upload_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_designation_upload.mt_alert_designation_upload_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }
    
    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        _logger.info("_needaction_domain_get - %s", users_obj)

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            _logger.info("_needaction_domain_get - is HR Manager")
            domain = [('state', '=', 'confirm')]
            return domain

        return False

    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Earning/Deduction Upload action!'),
                    _('Earning/Deduction Upload action has been initiated. Either cancel the designation_upload action or create another to undo it.')
                )

        return super(ng_state_payroll_designation_upload, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for o in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                o.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def designation_upload_state_confirm(self, cr, uid, ids, context=None):
        _logger.info("before state_confirm - %d", uid)
        self.write(cr, uid, ids, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def try_confirmed_designation_upload_actions(self, cr, uid, context=None):
        _logger.info("Running try_confirmed_designation_upload_actions cron-job...")
        employee_obj = self.pool.get('hr.employee')
        upload_obj = self.pool.get('ng.state.payroll.designation.upload')
        today = datetime.now().date()

        cr.execute('deallocate all')
        cr.execute('prepare update_designation (int, int) as update hr_employee set designation_id = $1 where id = $2')

        upload_ids = upload_obj.search(cr, uid, [('state', '=', 'confirm')], context=context)
        
        for upload in self.browse(cr, uid, upload_ids, context=context):
            if upload.upload_file != None and datetime.strptime(upload.date, DEFAULT_SERVER_DATE_FORMAT).date() <= today and upload.state == 'confirm':
                exception_list = []
                data_file = base64.decodestring(upload.upload_file)
                wb = open_workbook(file_contents=data_file)
                for s in wb.sheets():
                    _logger.info("Number of sheets: %d", len(wb.sheets()))
                    _logger.info("Number of records: %d", s.nrows)
                    for row in range(s.nrows):
                        if row > 0: #Skip first row
                            data_row = []
                            for col in range(s.ncols):
                                value = (s.cell(row, col).value)
                                data_row.append(value)
                            cr.execute("select id from hr_employee where employee_no='" + str(data_row[0]).strip().replace("'", "") + "'")
                            employee_id = cr.fetchall()                            
                            if employee_id and len(employee_id) == 1 and len(data_row) == 2:
                                designation = str(data_row[1]).strip().replace("'", "")
                                cr.execute("select id from ng_state_payroll_designation where code='" + designation + "'")
                                designation_id = cr.fetchall()
                                if designation_id:
                                    cr.execute('execute update_designation(%s,%s)', (designation_id[0],employee_id[0]))
                                else:
                                    exception_list.append({'employee_no':data_row[0],'designation':data_row[1],'error':'Designation with code not found; - ' + designation}) 
                            else:
                                _logger.info("%s - Employee Length: %d, Data Length: %d", data_row[0],len(employee_id),len(data_row))
                                if len(employee_id) != 1:
                                    exception_list.append({'employee_no':data_row[0],'designation':'','error':'No employee found for Employee Number'}) 
                                else:
                                    exception_list.append({'employee_no':data_row[0],'designation':'','error':'Wrong number of spreadsheet columns'})
                cr.execute("update ng_state_payroll_designation_upload set state='approved' where id=" + str(upload.id))                
                if len(exception_list) > 0:
                    with open('/tmp/' + 'designation_upload_exceptions_' + str(upload.id) + '.csv', 'w') as csvfile:
                        fieldnames = ['employee_no', 'designation','error']
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        writer.writeheader()
                        writer.writerows(exception_list)
                        csvfile.close()

                message = "Dear Sir/Madam,\nUpload of Designation file has completed.\n\nThank you.\n"
                message_exception = "\nPS: There were " + str(len(exception_list)) + " exceptions."
                smtp_obj = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
                smtp_obj.ehlo()
                #smtp_obj.starttls()
                #smtp_obj.ehlo()
                smtp_obj.login(user="osun.payroll@gmail.com", password="p@55w0rd1939")
                sender = 'osun.payroll@gmail.com'
                receivers = upload.notify_emails #Comma separated email addresses
                msg = MIMEMultipart()
                msg['Subject'] = 'Upload Completed' 
                msg['From'] = sender
                #msg['To'] = ', '.join(receivers)
                msg['To'] = receivers
                                 
                part = False
                if len(exception_list) > 0:
                    part = MIMEBase('application', "octet-stream")
                    part.set_payload(open('/tmp/' + 'designation_upload_exceptions_' + str(upload.id) + '.csv', "rb").read())
                    Encoders.encode_base64(part)                            
                    part.add_header('Content-Disposition', 'attachment; filename="designation_upload_exceptions_' + str(upload.id) + '.csv"')
                    message = message + message_exception
                    msg.attach(MIMEText(message))
                                
                    if part:
                        msg.attach(part)
                                                        
                smtp_obj.sendmail(sender, receivers, msg.as_string())         
                _logger.info("Email successfully sent to: " + receivers)   
            else:
                return False
        cr.commit()

        return True
    
#                 
# class ng_state_payroll_dashboard_item(models.Model):
#     '''
#     Payroll Dashboard Item
#     '''
#     _name = "ng.state.payroll.dashboard.item"
#     _description = 'Payroll Dashboard Item'
#     
#     _columns = {
#         'dashboard_id': fields.many2one('ng.state.payroll.dashboard', 'Payroll Dashboard', required=True),
#         'department_id': fields.many2one('hr.department', 'MDA', required=True),
#         'expected_gross': fields.float('Expected Gross', help='Expected Gross'),
#         'actual_gross': fields.float('Actual Gross', help='Actual Gross'),
#     }
#             
# class ng_state_payroll_dashboard(models.Model):
#     '''
#     Payroll Dashboard
#     '''
#     _name = "ng.state.payroll.dashboard"
#     _description = 'Payroll Dashboard'
#     
#     _columns = {
#         'payroll_id': fields.many2one('ng.state.payroll.payroll', 'Payroll', required=True),
#         'item_ids': fields.one2many('ng.state.payroll.dashboard.item','dashboard_id','Dashboard Items'),
#     }
#     
#     @api.multi
#     def name_get(self):
#         data = []
#         for d in self:
#             display_value = ''
#             display_value += d.payroll_id.name
#             display_value += '_'
#             display_value += str(d.id)
#             data.append((d.id, display_value))
#             
#         return data
#             
#     @api.model
#     def create(self, vals):
#         payroll_id = vals['payroll_id']
#         _logger.info("Creating dashboard for payroll - %d...", payroll_id.id)
#         
#         employees = False
#         if not payroll_id.create_user.domain_mdas:
#             _logger.info("No Domain MDAs.")
#             employees = self.env['hr.employee'].search([('resolved_earn_dedt', '=', True), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id')
#         else:
#             _logger.info("Domain MDAs= %s", payroll_id.create_user.domain_mdas)
#             employees = self.env['hr.employee'].search([('resolved_earn_dedt', '=', True), ('department_id.id', 'in', payroll_id.create_user.domain_mdas.ids), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id')
# 
#         _logger.info("Count employees= %d", len(employees))
#         
#         dashboard_id = False
#         if employees and len(employees) > 0:
#             department_ids = Set()
#             for emp in employees:
#                 if emp.department_id.id:
#                     department_ids.add(emp.department_id.id)
#     
#             self.env.cr.execute('insert into ng_state_payroll_dashboard (payroll_id) values (' + str(payroll_id.id) + ') returning id')
#             dashboard_id = self.env.cr.fetchone()
#             #self.env.cr.execute('deallocate all');
#             #self.env.cr.execute('prepare insert_dashboard_item (int, int) as insert into ng_state_payroll_dashboard_item (dashboard_id,department_id) values ($1,$2)')
#                     
#             for dept_id in department_ids:
#                 self.env.cr.execute('insert into ng_state_payroll_dashboard_item (dashboard_id,department_id) values (' + str(dashboard_id[0]) + ',' + str(dept_id) + ')')
#                 #self.env.cr.execute('execute insert_dashboard_item(%s,%s)', (dashboard_id.id,dept_id))
#             self.env.cr.commit()
#         
#         return dashboard_id
                
class ng_state_payroll_diff_detail_item(models.Model):
    '''
    Materiality Difference Detail Item
    '''
    _name = "ng.state.payroll.diff.detail.item"
    _description = 'Materiality Difference Detail Item'
    
    _columns = {
        'diff_summary_id': fields.many2one('ng.state.payroll.diff.summary.item', 'Materiality Difference Summary Item', required=True),
        'source_employee_id': fields.many2one('hr.employee', 'Source Employee', required=False),
        'target_employee_id': fields.many2one('hr.employee', 'Target Employee', required=False),
        'source_gross': fields.float('Source Gross', help='Source Gross'),
        'target_gross': fields.float('Target Gross', help='Target Gross'),
        'difference': fields.float('Gross Difference', help='Gross Difference'),
    }
                
class ng_state_payroll_diff_detail2_item(models.Model):
    '''
    Materiality Difference Detail2 Item
    '''
    _name = "ng.state.payroll.diff.detail2.item"
    _description = 'Materiality Difference Detail2 Item'
    
    _columns = {
        'diff_summary_id': fields.many2one('ng.state.payroll.diff.summary2.item', 'Materiality Difference Summary Item', required=True),
        'source_employee_id': fields.many2one('hr.employee', 'Source Employee', required=False),
        'target_employee_id': fields.many2one('hr.employee', 'Target Employee', required=False),
        'source_gross': fields.float('Source Gross', help='Source Gross'),
        'target_gross': fields.float('Target Gross', help='Target Gross'),
        'difference': fields.float('Gross Difference', help='Gross Difference'),
    }
                
class ng_state_payroll_diff_summary_item(models.Model):
    '''
    Materiality Difference Summary Item (MDA-based)
    '''
    _name = "ng.state.payroll.diff.summary.item"
    _description = 'Materiality Difference Summary Item'
    
    _columns = {
        'diff_id': fields.many2one('ng.state.payroll.diff', 'Materiality Difference', required=True),
        'source_department_id': fields.many2one('hr.department', 'Source MDA', required=False),
        'target_department_id': fields.many2one('hr.department', 'Target MDA', required=False),
        'source_gross': fields.float('Source Gross', help='Source Gross'),
        'target_gross': fields.float('Target Gross', help='Target Gross'),
        'difference': fields.float('Gross Difference', help='Gross Difference'),
        'detail_item_ids': fields.one2many('ng.state.payroll.diff.detail.item','diff_summary_id','Items'),
    }
                
class ng_state_payroll_diff_summary2_item(models.Model):
    '''
    Materiality Difference Summary Item (TCO-based)
    '''
    _name = "ng.state.payroll.diff.summary2.item"
    _description = 'Materiality Difference Summary Item'
    
    _columns = {
        'diff_id': fields.many2one('ng.state.payroll.diff', 'Materiality Difference', required=True),
        'source_tco_id': fields.many2one('ng.state.payroll.tco', 'Source TCO', required=False),
        'target_tco_id': fields.many2one('ng.state.payroll.tco', 'Target TCO', required=False),
        'source_gross': fields.float('Source Gross', help='Source Gross'),
        'target_gross': fields.float('Target Gross', help='Target Gross'),
        'difference': fields.float('Gross Difference', help='Gross Difference'),
        'detail_item_ids': fields.one2many('ng.state.payroll.diff.detail.item','diff_summary_id','Items'),
    }
                
class ng_state_payroll_diff_subvention_item(models.Model):
    '''
    Materiality Difference Subvention Item
    '''
    _name = "ng.state.payroll.diff.subvention.item"
    _description = 'Materiality Difference Subvention Item'
    
    _columns = {
        'diff_id': fields.many2one('ng.state.payroll.diff', 'Materiality Difference', required=True),
        'source_department_id': fields.many2one('hr.department', 'Source MDA', required=False),
        'target_department_id': fields.many2one('hr.department', 'Target MDA', required=False),
        'source_amount': fields.float('Source Amount', help='Source Amount'),
        'target_amount': fields.float('Target Amount', help='Target Amount'),
        'difference': fields.float('Amount Difference', help='Amount Difference'),
    }
            
class ng_state_payroll_diff(models.Model):
    '''
    Materiality Difference
    '''
    _name = "ng.state.payroll.diff"
    _description = 'Materiality Difference'
    
    _columns = {
        'detailed_diff': fields.boolean('Detailed Diff', help='Perform a granular comparison'),
        'source_payroll_id': fields.many2one('ng.state.payroll.payroll', 'Source Payroll', required=True),
        'target_payroll_id': fields.many2one('ng.state.payroll.payroll', 'Target Payroll', required=True),
        'summary_item_ids': fields.one2many('ng.state.payroll.diff.summary.item','diff_id','MDA Items'),
        'summary_item2_ids': fields.one2many('ng.state.payroll.diff.summary2.item','diff_id','TCO Items'),
        'subvention_item_ids': fields.one2many('ng.state.payroll.diff.subvention.item','diff_id','Items'),
        'notify_emails': fields.char('Notify Email', help='Comma separated email recipients for event notification', required=False),
        'state': fields.selection([
            ('draft','Draft'),
            ('pending','Pending'),
            ('processed','Processed'),
        ], 'Status'),        
    }
    
    _defaults = {
        'state': 'draft',
    } 
        
    @api.model
    def create(self, vals):
        vals['state'] = 'draft'
        payroll_singleton = self.env['ng.state.payroll.payroll']
        source_payroll = payroll_singleton,browse(vals.get('source_payroll_id'))

        if source_payroll.do_payroll:
            vals['is_payroll'] = True
        else:
            vals['is_payroll'] = False

        res = super(ng_state_payroll_payroll, self).create(vals)
        
        return res
    
    @api.multi
    def name_get(self):
 
        data = []
        for d in self:
            display_value = ''
            display_value += d.source_payroll_id.name
            display_value += ' to '
            display_value += d.target_payroll_id.name
            display_value += '_'
            display_value += str(d.id)
            data.append((d.id, display_value))
            
        return data
              
    @api.multi
    def set_pending(self, context=None):
        _logger.info("Calling set_pending...")
        
        self.write({'state': 'pending'})
        return True  
        
    def try_process(self, cr, uid, context=None):
        _logger.info("Running diff cron-job...")
        diff_singleton = self.pool.get('ng.state.payroll.diff')
        diff_ids = diff_singleton.search(cr, uid, [('state', '=', 'pending')], context=context)
        diff_obj = None
        for diff_id in diff_ids:
            diff_obj = diff_singleton.browse(cr, uid, diff_id, context=context)
            diff_obj.summary_compare()

        return True
            
    @api.multi
    def summary_compare(self):
        _logger.info("Calling summary_compare.. id=%d", self.id)
        
        if self.source_payroll_id.do_payroll and self.target_payroll_id.do_payroll:
            diff_summary_items = []
            #TODO Create combined set of source and target ids
            consolidated_ids = Set()
            for summary_id in self.source_payroll_id.payroll_summary_ids:
                consolidated_ids.add(summary_id.department_id.id)
            for summary_id in self.target_payroll_id.payroll_summary_ids:
                consolidated_ids.add(summary_id.department_id.id)

            for dept_id in consolidated_ids:
                source_summary_id = self.source_payroll_id.payroll_summary_ids.filtered(lambda r: r.department_id.id == dept_id)
                target_summary_id = self.target_payroll_id.payroll_summary_ids.filtered(lambda r: r.department_id.id == dept_id)
                source_gross = 0
                target_gross = 0
                if source_summary_id:
                    source_gross = source_summary_id[0].total_gross_income
                if target_summary_id:
                    target_gross = target_summary_id[0].total_gross_income
                diff = source_gross - target_gross
                detail_item_ids = []
                if self.detailed_diff:
                    detail_item_ids = self.detail_compare(dept_id)
                diff_summary_items.append({
                    'diff_id':self.id,
                    'source_department_id':dept_id,
                    'target_department_id':dept_id,
                    'source_gross':source_gross,
                    'target_gross':target_gross,
                    'difference':diff,
                    'detail_item_ids':[(0, 0, x) for x in detail_item_ids]
                })
            self.write({'summary_item_ids':[(0, 0, x) for x in diff_summary_items], 'state': 'processed'})

            diff_subvention_items = []
            consolidated_ids = Set()
            for subvention_id in self.source_payroll_id.subvention_item_ids:
                consolidated_ids.add(subvention_id.department_id.id)
            for subvention_id in self.target_payroll_id.subvention_item_ids:
                consolidated_ids.add(subvention_id.department_id.id)

            for dept_id in consolidated_ids:
                source_subvention_id = self.source_payroll_id.subvention_item_ids.filtered(lambda r: r.department_id.id == dept_id)
                target_subvention_id = self.target_payroll_id.subvention_item_ids.filtered(lambda r: r.department_id.id == dept_id)
                source_amount = 0
                target_amount = 0
                if source_subvention_id:
                    source_amount = source_subvention_id[0].total_gross_income
                if target_subvention_id:
                    target_amount = target_subvention_id[0].total_gross_income
                diff = source_amount - target_amount
                detail_item_ids = []
                diff_subvention_items.append({
                    'diff_id':self.id,
                    'source_department_id':dept_id,
                    'target_department_id':dept_id,
                    'source_gross':source_amount,
                    'target_gross':target_amount,
                    'difference':diff
                })
            self.write({'subvention_item_ids':[(0, 0, x) for x in diff_subvention_items]})

            if self.notify_emails:
                message = "Dear Sir/Madam,\nProcessing of Materiality Difference has completed.\n\nThank you.\n"
                smtp_obj = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
                smtp_obj.ehlo()
                #smtp_obj.starttls()
                #smtp_obj.ehlo()
                smtp_obj.login(user="osun.payroll@gmail.com", password="p@55w0rd1939")
                sender = 'osun.payroll@gmail.com'
                receivers = self.notify_emails #Comma separated email addresses
                msg = MIMEMultipart()
                msg['Subject'] = 'Materiality Difference Complete' 
                msg['From'] = sender
                #msg['To'] = ', '.join(receivers)
                msg['To'] = receivers
                msg.attach(MIMEText(message))                        
                smtp_obj.sendmail(sender, receivers, msg.as_string())         
                _logger.info("Email successfully sent to: " + receivers)
        if self.source_payroll_id.do_pension and self.target_payroll_id.do_pension:
            diff_summary_items = []
            #TODO Create combined set of source and target ids
            consolidated_ids = Set()
            for summary_id in self.source_payroll_id.pension_summary_ids:
                consolidated_ids.add(summary_id.tco_id.id)
            for summary_id in self.target_payroll_id.pension_summary_ids:
                consolidated_ids.add(summary_id.tco_id.id)

            for tco_id in consolidated_ids:
                source_summary_id = self.source_payroll_id.pension_summary_ids.filtered(lambda r: r.tco_id.id == tco_id)
                target_summary_id = self.target_payroll_id.pension_summary_ids.filtered(lambda r: r.tco_id.id == tco_id)
                source_gross = 0
                target_gross = 0
                if source_summary_id:
                    source_gross = source_summary_id[0].total_gross_income
                if target_summary_id:
                    target_gross = target_summary_id[0].total_gross_income
                diff = source_gross - target_gross
                detail_item_ids = []
                if self.detailed_diff:
                    detail_item_ids = self.detail_compare(tco_id)
                diff_summary_items.append({
                    'diff_id':self.id,
                    'source_tco_id':tco_id,
                    'target_tco_id':tco_id,
                    'source_gross':source_gross,
                    'target_gross':target_gross,
                    'difference':diff,
                    'detail_item_ids':[(0, 0, x) for x in detail_item_ids]    
                })
            self.write({'summary_item2_ids':[(0, 0, x) for x in diff_summary_items], 'state': 'processed'})
            if self.notify_emails:
                message = "Dear Sir/Madam,\nProcessing of Materiality Difference has completed.\n\nThank you.\n"
                smtp_obj = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
                smtp_obj.ehlo()
                #smtp_obj.starttls()
                #smtp_obj.ehlo()
                smtp_obj.login(user="osun.payroll@gmail.com", password="p@55w0rd1939")
                sender = 'osun.payroll@gmail.com'
                receivers = self.notify_emails #Comma separated email addresses
                msg = MIMEMultipart()
                msg['Subject'] = 'Materiality Difference Complete' 
                msg['From'] = sender
                #msg['To'] = ', '.join(receivers)
                msg['To'] = receivers
                msg.attach(MIMEText(message))                        
                smtp_obj.sendmail(sender, receivers, msg.as_string())         
                _logger.info("Email successfully sent to: " + receivers)      
    
    @api.multi
    def detail_compare(self, dept_id, is_payroll):
        diff_detail_items = []
        consolidated_ids = Set()
        source_item_ids = []
        target_item_ids = []
        if is_payroll:
	    source_item_ids = self.source_payroll_id.payroll_item_ids.filtered(lambda r: r.employee_id.department_id.id == dept_id and r.active)
            target_item_ids = self.target_payroll_id.payroll_item_ids.filtered(lambda r: r.employee_id.department_id.id == dept_id and r.active)
        else:
	    source_item_ids = self.source_payroll_id.pension_item_ids.filtered(lambda r: r.employee_id.tco_id.id == dept_id and r.active)
            target_item_ids = self.target_payroll_id.pension_item_ids.filtered(lambda r: r.employee_id.tco_id.id == dept_id and r.active)

        for item_id in source_item_ids:
            consolidated_ids.add(item_id.employee_id.id)
        for item_id in target_item_ids:
            consolidated_ids.add(item_id.employee_id.id)
        for emp_id in consolidated_ids:    
            source_payroll_item_id = source_item_ids.filtered(lambda r: r.employee_id.id == emp_id)
            target_payroll_item_id = target_item_ids.filtered(lambda r: r.employee_id.id == emp_id)
            source_gross = 0
            target_gross = 0
            if source_payroll_item_id:
                source_gross = source_payroll_item_id[0].gross_income
            if target_payroll_item_id:
                target_gross = target_payroll_item_id[0].gross_income
            diff = source_gross - target_gross
            diff_rounded = Decimal(diff)
            diff_rounded = diff_rounded.quantize(Decimal('.01'), rounding=ROUND_DOWN)
            if not diff_rounded.is_zero():
                diff_detail_items.append({
                    'source_employee_id':emp_id,
                    'target_employee_id':emp_id,
                    'source_gross':source_gross,
                    'target_gross':target_gross,
                    'difference':diff
                })


        return diff_detail_items    

class ng_state_payroll_payroll(models.Model):
    '''
    Payroll
    '''
    _name = "ng.state.payroll.payroll"
    _description = 'Payroll'

    _inherit = ['mail.thread', 'ir.needaction_mixin']

    _columns = {
        'name': fields.char('Name', help='Payroll Name', required=True),
        'payroll_prev_id': fields.many2one('ng.state.payroll.payroll', 'Previous Month Payroll', required=False),
        'calendar_id': fields.many2one('ng.state.payroll.calendar', 'Calendar', track_visibility='onchange', required=True),
        'create_user': fields.many2one('res.users', 'Create User', required=True, readonly=1),
        'total_net_payroll': fields.float('Payroll Total Net', help='Payroll Total Net'),
        'total_gross_payroll': fields.float('Payroll Total Gross', help='Payroll Total Gross'),
        'total_taxable_payroll': fields.float('Payroll Total Taxable', help='Payroll Total Taxable'),
        'total_tax_payroll': fields.float('Payroll Total Tax', help='Payroll Total Tax'),
        'total_balance_payroll': fields.float('Payroll Total Balance', help='Total Balance Payroll Payment'),
        'total_net_pension': fields.float('Pension Total Net', help='Pension Total Net'),
        'total_gross_pension': fields.float('Pension Total Gross', help='Pension Total Gross'),
        'total_balance_pension': fields.float('Pension Total Balance', help='Total Balance Pension Payment'),
        'processing_time_payroll': fields.float('Payroll Processing Time', help='Payroll Processing Time'),
        'processing_time_pension': fields.float('Pension Processing Time', help='Pension Processing Time'),
        'notify_emails': fields.char('Notify Email', help='Comma separated email recipients for event notification', required=False),
        'mda_emails': fields.char('MDA Email', help='Comma separated email recipients for MDA notification', required=False),
        'from_date': fields.related('calendar_id', 'from_date', type='date', string='From Date', readonly=1),
        'to_date': fields.related('calendar_id', 'to_date', type='date', string='To Date', readonly=1),
        'payroll_item_ids': fields.one2many('ng.state.payroll.payroll.item','payroll_id','Payroll Items'),
        'pension_item_ids': fields.one2many('ng.state.payroll.pension.item','payroll_id','Pension Items'),
#         'dashboard_ids': fields.one2many('ng.state.payroll.dashboard','payroll_id','Dashboards'),
        'subvention_item_ids': fields.one2many('ng.state.payroll.subvention.item','payroll_id','Items'),
        'payroll_schoolsummary_ids': fields.one2many('ng.state.payroll.payroll.schoolsummary','payroll_id','Payroll School Summary Items'),
        'payroll_summary_ids': fields.one2many('ng.state.payroll.payroll.summary','payroll_id','Payroll Summary Items'),
        'pension_summary_ids': fields.one2many('ng.state.payroll.pension.summary','payroll_id','Pension Summary Items'),
        'signoff_ids': fields.one2many('ng.state.payroll.payroll.signoff','payroll_id','Sign-Off Items'),
        'signoff_pos_order': fields.integer('Sign-off Index', help='Sign-off Index'),
        'scenario_ids': fields.one2many('ng.state.payroll.scenario','payroll_id','Scenario Payments'),       
        'do_dry_run': fields.boolean('Dry Run', help='Tick check-box to do dry run'),
        'auto_process': fields.boolean('Auto-process', help='Tick check-box for automatic processing'),
        'in_progress': fields.boolean('In Progress', help='Indicates processing currently in progress'),
        'do_payroll': fields.boolean('Run Payroll', help='Tick check-box to run active employee payroll'),
        'do_pension': fields.boolean('Run Pension', help='Tick check-box to run pension payroll'),
        'apply_ta_rules': fields.boolean('T&A', help='Activate T&A rules'),
        'generate_payslips': fields.boolean('Generate Payslips', help='Generate Payslips'),
        'generate_reports': fields.boolean('Generate Reports', help='Generate Reports'),
        'gov_sign': fields.binary('Governor Signature'),
        'ps_finance_sign': fields.binary('PS Finance Signature'),
        'payroll_report': fields.binary('Payroll Report'),
        'paye_report': fields.binary('PAYE Report'),
        'summary_report': fields.binary('Payroll Summary Report'),
        'pension_exec_summary_report': fields.binary('Pension Executive Summary Report'),
        'exec_summary_report': fields.binary('Full Wagebill Summary - MDA/TESCOM/LTH'),
        'exec_summary2_report': fields.binary('Full Wagebill Summary - SUBEB/LGA'),
        'pension_report': fields.binary('Pension Report'),
        'pension_mda_report': fields.binary('Pension Report'),
        'departments_report': fields.binary('Departments Report'),
        'tescom_report': fields.binary('TESCOM Report'),
        'tescom_school_report': fields.binary('TESCOM School Report'),
        'mda_report': fields.binary('MDA Report'),
        'tco_report': fields.binary('TCO Report'),
        'mda_deduction_report': fields.binary('MDA Deduction Report'),
        'mda_deduction_head_report': fields.binary('MDA Deduction Head Report'),
        'mda_summary_report': fields.binary('MDA Summary Report'),
        'leavebonus_report': fields.binary('Leave Allowance Report'),
        'state': fields.selection([
            ('draft','Draft'),
            ('pending','Pending'),
            ('in_progress','Processing'),
            ('processed','Processed'),
            ('closed','Closed'),
        ], 'Status')
    }
    
    _defaults = {
        'state': 'draft',
        'signoff_pos_order': 0,
        'do_dry_run': False,
        'auto_process': True,
        'generate_payslips': True,
        'generate_reports': True,
    }                                          

    _track = {
        'state': {
            'ng_state_payroll_payroll.mt_alert_promo_pending':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'pending',
            'ng_state_payroll_payroll.mt_alert_promo_in_progress':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'in_progress',
            'ng_state_payroll_payroll.mt_alert_promo_processed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'processed',
            'ng_state_payroll_payroll.mt_alert_promo_closed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'closed',
        },
    }

    def try_generate_mda_payslips(self, cr, uid, context=None):
        """Generate list of payslips for each. Called from
        the scheduler."""

        _logger.info("Running try_generate_mda_payslips cron-job...")
        
        payroll_obj = self.pool.get('ng.state.payroll.payroll')
        payroll_ids = payroll_obj.search(cr, uid, [
            ('state', '=', 'closed'),
            ('generate_payslips', '=', True),
        ], context=context)
                
        template_doc = DocxTemplate("/odoo/odoo9/odoo-server/addons/hr_payroll_nigerian_state/payslip_tpl.docx")

        for p in self.browse(cr, uid, payroll_ids, context=context):
            _logger.info("Generating Payslips for: %s", p.name)
            for p_item in p.payroll_item_ids.filtered(lambda r: r.active):
                payslip_dir = "/tmp/payslips/" + str(p.id) + "/" + str(p_item.employee_id.department_id.id)
                if not os.path.exists(payslip_dir):
                    os.makedirs(payslip_dir)
                earnings_std = p_item.item_line_ids.filtered(lambda r: 'OTHER EARNINGS' not in r.name and r.amount >= 0)
                earnings_nstd = p_item.item_line_ids.filtered(lambda r: 'OTHER EARNINGS' in r.name and r.amount >= 0)
                deductions_std = p_item.item_line_ids.filtered(lambda r: 'OTHER DEDUCTIONS' not in r.name and r.amount < 0)
                deductions_nstd = p_item.item_line_ids.filtered(lambda r: 'OTHER DEDUCTIONS' in r.name and r.amount < 0)
                earnings_total = 0
                deductions_total = 0
                for p_item_line in earnings_std:
                    earnings_total += p_item_line.amount
                for p_item_line in earnings_nstd:
                    earnings_total += p_item_line.amount
                for p_item_line in deductions_std:
                    deductions_total += -p_item_line.amount
                for p_item_line in deductions_nstd:
                    deductions_total += -p_item_line.amount
                gross_income = p_item.gross_income
                net_income = p_item.net_income
                context = { 'emp':p_item.employee_id,
                           'calendar': p.calendar_id, 
                           'earnings_std':earnings_std, 
                           'earnings_nstd':earnings_nstd, 
                           'deductions_std':deductions_std, 
                           'deductions_nstd':deductions_nstd,
                           'earnings_total':earnings_total,
                           'deductions_total':deductions_total,
                           'gross_income':gross_income,
                           'net_income':net_income, 
                        }
                template_doc.render(context)
                template_doc.save(payslip_dir + "/payslip_" + str(p_item.employee_id.id) + ".docx")            
            p.write({'generate_payslips':False})

        return True
        
    @api.depends('signoff_ids')
    def _check_user_signer(self):
        _logger.info("Calling _check_user_signer..state = %s", self.state)
        self.current_user_signer = False
        #if self.env.user.groups_id.name == 'Payroll Officer':
        #    for sign_off in self.signoff_ids:
        #        if sign_off.user_id.id == self.env.user.id:
        #            self.current_user_signer = True
        #            break
   
    @api.multi
    def run_dry_run(self, vals):
        return self.dry_run()
        
    @api.multi
    def unlink(self):
        self.env.cr.execute("alter table ng_state_payroll_pension_item disable trigger all")
        self.env.cr.execute("alter table ng_state_payroll_payroll_item disable trigger all")
        for p_obj in self:
            self.env.cr.execute("delete from ng_state_payroll_scenario where payroll_id=" + str(p_obj.id))
    
            if p_obj.do_payroll:
                self.env.cr.execute("delete from ng_state_payroll_payroll_item_line where item_id in (select id from ng_state_payroll_payroll_item where payroll_id=" + str(p_obj.id) + ")")
                self.env.cr.execute("delete from ng_state_payroll_payroll_item where payroll_id=" + str(p_obj.id))
                self.env.cr.execute("delete from ng_state_payroll_subvention_item where payroll_id=" + str(p_obj.id))
                self.env.cr.execute("delete from ng_state_payroll_payroll_summary where payroll_id=" + str(p_obj.id))
            if p_obj.do_pension:
                self.env.cr.execute("delete from ng_state_payroll_pension_item_line where item_id in (select id from ng_state_payroll_pension_item where payroll_id=" + str(p_obj.id) + ")")
                self.env.cr.execute("delete from ng_state_payroll_pension_item where payroll_id=" + str(p_obj.id))
            
#             self.env.cr.execute("delete from ng_state_payroll_dashboard_item where dashboard_id in (select id from ng_state_payroll_dashboard where payroll_id=" + str(p_obj.id) + ")")
#             self.env.cr.execute("delete from ng_state_payroll_dashboard where payroll_id=" + str(p_obj.id))
            self.env.cr.execute("delete from ng_state_payroll_payroll where id=" + str(p_obj.id))
        self.env.cr.execute("alter table ng_state_payroll_pension_item enable trigger all")
        self.env.cr.execute("alter table ng_state_payroll_payroll_item enable trigger all")
        self.env.invalidate_all()
        
    @api.model
    def create(self, vals):
        vals['state'] = 'draft'
        vals['create_user'] = self.env.user.id
        if vals.has_key('auto_process') and not vals['auto_process']:
            vals['state'] = 'in_progress'
        res = super(ng_state_payroll_payroll, self).create(vals)
        
#         if vals['do_payroll']:
#             self.env['ng.state.payroll.dashboard'].create({'payroll_id': res})
        
        return res

    @api.multi
    def write(self, vals):
        _logger.info("Calling write.. id= %d", self.id)
        if vals.has_key('auto_process') and not vals['auto_process']:
            if vals['state'] == 'draft':
                vals['state'] = 'in_progress'
        if vals.has_key('auto_process') and vals['auto_process']:
            if vals.has_key('state') and vals['state'] == 'in_progress':
                vals['state'] = 'draft'
                
        return super(ng_state_payroll_payroll,self).write(vals)

    def list_payroll_items(self, cr, uid, context=None):
        _logger.info("Calling list_payroll_items")
        _logger.info("User ID=%s", uid)
        employee_obj = self.pool.get('hr.employee')
        employee_id = employee_obj.search(cr, uid, [('user_id', '=', uid)])
        emp = employee_obj.browse(cr, uid, employee_id, context=context)
        payroll_items = []
        if employee_id:
            _logger.info("Employee=%d[%s]", employee_id[0], emp.name)
            cr.execute("select id,(select name from ng_state_payroll_payroll where id=pitem.payroll_id) from ng_state_payroll_payroll_item as pitem where employee_id=" + str(employee_id[0]))
            payroll_items = cr.fetchall()
        else:
            _logger.info("No matching employee found for ID %d", uid)
        
        _logger.info("Items=%s", payroll_items)   
        return payroll_items
     
    @api.multi
    def reset_reports(self, vals):
        _logger.info("Calling reset_reports..vals = %s", vals)
        if self.do_payroll:
            self.env.cr.execute("update ng_state_payroll_payroll set payroll_report=null,summary_report=null,departments_report=null,mda_report=null,tescom_report=null,tescom_school_report=null,mda_deduction_report=null,mda_summary_report=null,exec_summary_report=null,exec_summary2_report=null,leavebonus_report=null where id=" + str(self.id))
        
        if self.do_pension:
            self.env.cr.execute("update ng_state_payroll_payroll set pension_report=null,pension_mda_report=null where id=" + str(self.id))
        self.env.invalidate_all()
           
    @api.multi
    def revert(self):
        _logger.info("Calling revert..id = %s", self.id)
        #if self.env.user.has_group('hr_payroll_nigerian_state.group_payroll_admin'):

        if self.do_payroll:
            self.env.cr.execute("alter table ng_state_payroll_payroll_item disable trigger all")
            self.env.cr.execute("update ng_state_payroll_payroll set total_tax_payroll=0,total_net_payroll=0,total_gross_payroll=0,total_taxable_payroll=0,total_balance_payroll=0,processing_time_payroll=0,payroll_report=null,paye_report=null,summary_report=null,exec_summary_report=null,exec_summary2_report=null,departments_report=null,tescom_report=null,tescom_school_report=null,mda_report=null,mda_deduction_report=null,mda_deduction_head_report=null,mda_summary_report=null,leavebonus_report=null,state='draft' where id=" + str(self.id))
            self.env.cr.execute("delete from ng_state_payroll_pension_item_line where item_id in (select id from ng_state_payroll_payroll_item where payroll_id=" + str(self.id) + ")")
            self.env.cr.execute("delete from ng_state_payroll_payroll_item where payroll_id=" + str(self.id))
            self.env.cr.execute("delete from ng_state_payroll_payroll_schoolsummary where payroll_id=" + str(self.id))
            self.env.cr.execute("delete from ng_state_payroll_payroll_summary where payroll_id=" + str(self.id))
            self.env.cr.execute("alter table ng_state_payroll_payroll_item enable trigger all")
            self.env.cr.execute("update ng_state_payroll_payroll set payroll_report=null,summary_report=null,departments_report=null,mda_report=null,mda_deduction_report=null,mda_summary_report=null,exec_summary_report=null,exec_summary2_report=null,leavebonus_report=null where id=" + str(self.id))
            
        if self.do_pension:
            self.env.cr.execute("alter table ng_state_payroll_pension_item disable trigger all")
            self.env.cr.execute("update ng_state_payroll_payroll set total_net_pension=0,total_gross_pension=0,total_balance_pension=0,processing_time_pension=0,pension_exec_summary_report=null,pension_mda_report=null,pension_report=null,state='draft' where id=" + str(self.id))
            self.env.cr.execute("delete from ng_state_payroll_pension_item_line where item_id in (select id from ng_state_payroll_pension_item where payroll_id=" + str(self.id) + ")")
            self.env.cr.execute("delete from ng_state_payroll_pension_item where payroll_id=" + str(self.id))
            self.env.cr.execute("delete from ng_state_payroll_pension_summary where payroll_id=" + str(self.id))
            self.env.cr.execute("alter table ng_state_payroll_pension_item enable trigger all")
            self.env.cr.execute("update ng_state_payroll_payroll set pension_report=null,pension_mda_report=null where id=" + str(self.id))
        self.env.invalidate_all()

        files = os.listdir(REPORTS_DIR)
        for file in files:
            if file.endswith('_' + str(self.id) + '.xlsx'):
                os.remove(os.path.join(REPORTS_DIR, file))

        
    @api.multi
    def sign_off(self):        
        _logger.info("Calling sign_off..state = %s", self.state)
        #TODO Set sign-off entry for current user to true
        group_payroll_officer = self.env['res.groups'].search([('name', '=', 'Payroll Officer')])
        group_admin = self.env['res.groups'].search([('name', '=', 'Configuration')])
        #if group_payroll_officer in self.env.user.groups_id or group_admin in self.env.user.groups_id:
        if True:
            #Iterate through sign-off users and if all signed off, set state='closed'
            signoff_count = 0
            for sign_off in self.signoff_ids:
                if sign_off.user_id.id == self.env.user.id:
                    self.update({'signoff_pos_order': (self.signoff_pos_order + 1)})
                    sign_off.update({'signed_off': True})
                if sign_off.signed_off:
                    signoff_count += 1
            if len(self.signoff_ids) == signoff_count:
                self.state = 'closed'
                self.update({'state': 'closed'})        
                #If state_flag = 'closed' mark all non-permanent nonstandard deductions and earnings for the payroll calendar period as inactive.
                nonstd_earnings_nonperm = self.env['ng.state.payroll.earning.nonstd'].search([('active', '=', True), ('permanent', '=', False), ('calendars.id', '=', self.calendar_id.id)])
                for o in nonstd_earnings_nonperm:
                    o.update({'active': False})
                nonstd_deductions_nonperm = self.env['ng.state.payroll.deduction.nonstd'].search([('active', '=', True), ('permanent', '=', False), ('calendars.id', '=', self.calendar_id.id)])
                for o in nonstd_deductions_nonperm:
                    o.update({'active': False})
          
    @api.multi
    def set_pending(self, context=None):
        _logger.info("Calling set_pending...")
        
        self.write({'state': 'pending'})
        return True   
    
    @api.multi
    def set_in_progress(self, context=None):
        _logger.info("Calling set_in_progress...")
        
        self.write({'state': 'in_progress'})
        return True   
    
    @api.multi
    def set_finalized(self, context=None):
        _logger.info("Calling set_finalized...")
        
        self.write({'state': 'processed'})
        return True   
    
    def try_finalize(self, cr, uid, context=None):
        _logger.info("Running payroll cron-job...")
        payroll_singleton = self.pool.get('ng.state.payroll.payroll')
        payroll_ids = payroll_singleton.search(cr, uid, [('state', '=', 'pending'),('auto_process', '=', True)], context=context)
        if payroll_ids:
            payroll_obj = payroll_singleton.browse(cr, uid, payroll_ids[0], context=context)
            payroll_obj.set_in_progress(context=context)
            payroll_obj.finalize(context=context)

        return True
                        
    @api.multi
    def exec_finalize(self, context=None):
        self.set_in_progress(context=context)
        self.finalize(context=context)
                        
    @api.multi
    def refinalize(self, context=None):
        self.set_in_progress(context=context)
        self.revert(context=context)
        self.finalize(context=context)
    
    @api.model
    def summarize2(self, is_payroll):
        if is_payroll:
            _logger.info("summarize2 payroll count = %d", len(self.payroll_item_ids))
            dept_summary = {}
            for payroll_item in self.payroll_item_ids.filtered(lambda r: r.active):
                if not dept_summary.has_key(payroll_item.employee_id.department_id.id):
                    dept_summary[payroll_item.employee_id.department_id.id] = {'department_id':payroll_item.employee_id.department_id.id,
                                                                               'department':payroll_item.employee_id.department_id.name,
                                                                               'total_taxable_income':0,
                                                                               'total_gross_income':0,
                                                                               'total_net_income':0,
                                                                               'total_paye_tax':0,
                                                                               'total_nhf':0,
                                                                               'total_pension':0,
                                                                               'total_other_deductions':0,
                                                                               'total_leave_allowance':0,
                                                                               'total_strength':0}
    
                nhf = 0
                pension = 0
                other = 0
                for line_item in payroll_item.item_line_ids:
                    if line_item.name.upper().find('PENSION') >= 0:
                        pension += line_item.amount
                    elif line_item.name.upper().find('NHF') >= 0:
                        nhf += line_item.amount
                    else:
                        other += line_item.amount
                    
                dept_summary[payroll_item.employee_id.department_id.id]['total_taxable_income'] += payroll_item.taxable_income
                dept_summary[payroll_item.employee_id.department_id.id]['total_gross_income'] += payroll_item.gross_income
                dept_summary[payroll_item.employee_id.department_id.id]['total_net_income'] += payroll_item.net_income
                dept_summary[payroll_item.employee_id.department_id.id]['total_paye_tax'] += payroll_item.paye_tax
                dept_summary[payroll_item.employee_id.department_id.id]['total_leave_allowance'] += payroll_item.leave_allowance
                dept_summary[payroll_item.employee_id.department_id.id]['total_nhf'] += nhf
                dept_summary[payroll_item.employee_id.department_id.id]['total_pension'] += pension
                dept_summary[payroll_item.employee_id.department_id.id]['total_other_deductions'] += other
                dept_summary[payroll_item.employee_id.department_id.id]['total_strength'] += 1
            
            _logger.info("summarize2 payroll dict = %s", dept_summary)    
            return dept_summary
        else:
            _logger.info("summarize2 pension count = %d", len(self.pension_item_ids))
            tco_summary = {}
            for pension_item in self.pension_item_ids.filtered(lambda r: r.active):
                if not tco_summary.has_key(pension_item.employee_id.tco_id.id):
                    tco_summary[pension_item.employee_id.tco_id.id] = {'tco_id':pension_item.employee_id.tco_id.id,
                                                                       'tco':pension_item.employee_id.tco_id.name,
                                                                               'total_gross_income':0,
                                                                               'total_net_income':0,
                                                                               'total_arrears':0,
                                                                               'total_dues':0,
                                                                               'total_strength':0}
    
                arrears = 0
                dues = 0
                for line_item in pension_item.item_line_ids:
                    if line_item.name.upper().find('ARREARS') >= 0:
                        arrears += line_item.amount
                    elif line_item.name.upper().find('NUP') >= 0 or line_item.name.upper().find('HOS') >= 0:
                        dues += line_item.amount
                    
                tco_summary[pension_item.employee_id.tco_id.id]['total_net_income'] += pension_item.gross_income + dues
                tco_summary[pension_item.employee_id.tco_id.id]['total_gross_income'] += (pension_item.gross_income + arrears)
                tco_summary[pension_item.employee_id.tco_id.id]['total_arrears'] += arrears
                tco_summary[pension_item.employee_id.tco_id.id]['total_dues'] += dues
                tco_summary[pension_item.employee_id.tco_id.id]['total_strength'] += 1
            
            _logger.info("summarize2 pension dict = %s", tco_summary)    
            return tco_summary
                                
    @api.multi
    def finalize(self, context=None):
        _logger.info("Calling finalize...state = %s", self.state)
        
        if self.state == 'in_progress':
            if self.calendar_id:
                if self.do_payroll:
                    tic = time.time()
                    #item_list = []
                            
                    #List all subvention earnings for this calendar period
                    subventions = self.env['ng.state.payroll.subvention'].search([('active', '=', True), ('calendar_id', '=', self.calendar_id.id), ('org_id.id', 'in', self.create_user.domain_mdas.ids)])
            
                    #List all tax rules
                    paye_taxrules = self.env['ng.state.payroll.taxrule'].search([('active', '=', True)])
                    
                    ta_rules = self.env['hr.ta.rule'].search([('active', '=', True)], order='presence desc')
                    
                    #Fetch all active employees *TODO* (and non-suspended employees)
                    employees = False
                    employees_pending = False
                    if not self.create_user.domain_mdas:
                        _logger.info("No Domain MDAs.")
                        employees = self.env['hr.employee'].search([('resolved_earn_dedt', '=', True), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id')
                        employees_pending = self.env['hr.employee'].search([('resolved_earn_dedt', '=', False), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id')
                    else:
                        _logger.info("Domain MDAs= %s", self.create_user.domain_mdas)
                        employees = self.env['hr.employee'].search([('resolved_earn_dedt', '=', True), ('department_id.id', 'in', self.create_user.domain_mdas.ids), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id')
                        employees_pending = self.env['hr.employee'].search([('resolved_earn_dedt', '=', False), ('department_id.id', 'in', self.create_user.domain_mdas.ids), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id')
                    #employees = self.env['hr.employee'].search([('id', 'in', [6128,4170,15140,15150,15152,15086,15088,15090,15102,15104,15106,19778,15842,15998,18934,15474,15476,16396,18092,18106,25072,24806,109223,109224,109225,109226,4192,5614,7732,7734,6770,5468,4618,5372,7218,7034,13302,15948,15950,109230,109218,109219,109220,109221,10592,12474,13160,15092,15094,15108,15110,15112,15114,15116,15118,15120,15080,15082,15084,15122,15124,15126,15128,18584,19068,19296,19588,23580,23696,13526,15844,26686,88338,68330,68332,68334,68336,68338,68340,68342,68344,68346,68348,68350,68352,68354,68356,68358,68360,68362,68364,68366,68368,68370,68372,68374,68376,68378,68380,68382,68384,68386,68388,68390,68392,68394,68396,68398,68400,68574,68576,68580,68578,68582,68584,68588,68586,68590,68592,68596,68594,78328,78330,78332,78334,78336,78338,78340,78342,78350,78352,78354,78344,78346,78348,78356,78358,78360,78362,78364,78366,78368,78370,78372,78374,78376,78386,78388,78378,78380,78382,78384,78390,78392,78394,78396,78398,78400,78582,78584,78586,78588,78590,78592,78594,78596,78598,78600,78602,78604,78606,88332,88326,88330,88334,88336,88340,88342,88346,88348,88350,88352,88354,88356,88358,88458,88462,88464,88466,88468,95197,95589,96261,96481,97383,97403,97404,97405,97406,97407,97408,97409,97410,97423,97424,97425,97426,97836,102588,102630,108838,108927,109056])])
                    #employees = self.env['hr.employee'].search([('id', 'in', [5034,3736,3204,4838,2928,2932,3490,7430,7434,7436,6848,6060,5116,3232,2814,6682,7584,7310,7346,7350,13208,15416,15420,7488,14342,15234,6824,15244,15246,6872,6876,6880,7530,7534,7560,15252,8332,5366,14268,14270,14272,14274,20196,20198,20208,14348,14352,14424,13124,11218,20216,10724,10460,9680,10510,10926,17018,17020,20084,20086,17110,17116,19894,19896,20090,20092,20100,20104,20106,20944,19926,19516,21348,20146,20148,20156,20158,20412,20416,20454,20456,20460,20464,19102,19998,20002,19812,20326,20328,20046,20050,17230,19660,20288,20290,20292,20296,20300,20376,20380,20382,19766,19774,19780,19676,19824,19826,20304,20306,19788,19042,19046,19828,19830,19832,19834,19842,20666,19732,19734,19960,19970,19972,19974,17256,17260,17264,17268,17270,17272,17274,16500,15834,15840,15914,27114,16230,17336,17664,18932,16848,16584,17474,17478,18058,15454,16390,17824,17708,16668,16696,17624,17754,17730,17732,17388,17396,27980,27984,26714,26720,26964,27012,26744,16062,16910,16934,17346,17558,17990,27046,27048,27050,27052,22946,23800,24878,24880,27274,27318,27320,27278,27280,27288,26634,26638,27322,27324,26042,27302,24988,26052,24814,27408,23282,26128,26236,23080,87482,3206,3752,7460,5596,7462,7516,2462,5120,66896,6564,6782,4596,2874,4528,3212,3718,2760,3960,4958,4834,2672,4048,4228,4680,4012,4416,5272,5598,7354,7360,7364,7366,5656,7410,7414,7440,7444,7632,8364,5572,5306,6752,7464,7466,5790,7478,7480,7482,7486,7454,7456,7458,7678,7470,7662,7666,7932,6766,4554,4558,4312,6296,7694,7518,7526,5858,6728,6736,6744,7494,7498,7500,7502,7506,7510,7708,7714,5922,7050,7082,7528,7546,7554,6918,5200,5354,5358,8224,5958,12544,19692,15870,15978,26684,16022,7636,93997,2486,8510,16956,8846,8898,9364,6516,10710,10806,11002,11062,11108,11184,11260,11352,11468,11760,12108,12274,12578,12584,12890,13308,13412,13628,14122,14130,14250,14256,14340,14454,15230,15410,15440,16256,16300,16362,16430,16506,16636,16638,16704,16706,16716,16728,16834,16858,16884,16990,16960,16962,16970,16978,16980,16984,16994,17002,17004,17006,17038,17052,17054,17056,17076,17090,17124,17140,17142,17170,17176,17224,17226,17242,17290,17292,17294,17326,17332,17378,17380,17406,17412,17414,17416,17436,17504,17510,17512,17530,17572,17578,17596,17602,17644,17678,17694,17740,17784,17862,17834,17868,17950,18012,18014,18006,18066,18744,18718,18756,18758,18768,18830,18806,18856,18876,19088,19648,19636,19686,19690,19700,19736,19702,19704,19706,19708,19712,19718,19720,19722,19726,19728,19748,19750,19754,19790,19796,19802,19808,19846,19850,19852,19854,19816,19818,19820,19900,19908,19912,19918,19932,19934,19940,19942,19944,19946,19950,19956,19958,19978,20014,19994,19996,20018,20022,20028,20056,20036,20058,20064,20068,20076,20080,20108,20110,20112,20114,20116,20118,20120,20124,20094,20132,20174,20178,20182,20184,20186,20188,20194,20228,20238,20240,20242,20246,20250,20218,20220,20260,20262,20270,20286,20336,20342,20318,20322,20386,20348,20350,20360,20370,20426,20396,20400,20422,20446,20516,20518,20520,20474,20482,20484,20614,20670,20988,21144,21164,21304,22932,22960,23062,23446,23410,23468,23472,23570,23578,23624,23688,10038,24760,24768,24772,24784,24862,17276,24892,24890,24884,24944,24978,25024,19650,20514,20346,15848,19954,26062,26144,26158,26444,26590,26592,26598,26604,26606,26612,26616,26632,26688,26690,26678,26680,26682,26646,26692,26694,26704,26772,26956,26984,26992,27070,27222,27228,27290,27312,27354,27386,27436,27550,27612,27742,27868,27882,27884,28012,28112,28120,28826,29506,30392,31130,31300,31864,32234,32734,33642,34104,34438,34816,35362,36046,37038,37418,37996,38138,38246,38366,39770,40018,40390,40618,41324,42192,43120,43482,43486,43500,43512,43520,43516,43530,43532,43538,43544,43546,43570,43578,43584,43742,44596,44882,44944,45380,45400,45478,46278,46542,46654,46726,47646,47800,48336,48360,48604,49002,49024,49026,49028,49134,49166,49426,49454,49458,49464,49466,49472,50180,51324,51990,53964,56014,56168,56252,56680,57094,57118,57196,57202,57330,57360,57518,57538,57746,58370,58652,58692,60328,61044,61046,61776,61802,61922,64096,64440,65252,65332,65334,65326,65690,65824,65870,65956,66694,67496,67780,67902,75920,75926,76112,76210,76222,76232,76230,76256,76272,76962,76952,76960,77054,77556,77596,77600,77604,77598,77602,77610,77656,77660,80736,81008,81000,81002,81006,81014,81016,81018,81020,81024,83512,83514,83516,83518,83528,83530,83552,83698,83706,85968,85970,86068,86082,86090,86092,87014,87016,87018,87020,87022,87024,87354,87356,87498,87500,87558,89892,90194,90242,90244,90254,94498,93938,93940,93960,93965,93967,93968,93971,93973,93974,93998,93999,94017,94018,94032,94244,94246,94302,94303,94305,94306,94307,94363,94364,94365,94369,94403,94404,94419,94446,94448,94452,94497,94700,94848,94900,94903,94935,94937,94938,94944,94960,95135,95139,95343,95349,95351,95360,95357,95364,95365,95401,95637,95638,95641,95645,95649,95650,95651,95654,95655,95656,95660,95693,95726,95756,96293,96298,96321,96584,97163,97223,97224,97225,97226,97343,97795,102504,104487,108696,108734,109136])])
                    #employees = self.env['hr.employee'].search([('employee_no', 'in', ['27234','26707','24816','20251','27138','18833','19418','58922','25340','22675','18439','22621','24833','18910','27135','25116','21183','20075','58148','27699','21979'])])

                    _logger.info("Count employees= %d", len(employees))        
                    if employees_pending:
                        raise osv.except_osv(_('Error'), _('Employees requiring Pay Scheme/Grade resolution exist; wait 10 minutes for resolutions completion.'))
                    else:
                        exception_list = []
                        self.env.cr.execute('prepare insert_item (int,bool,int,numeric,numeric,numeric,numeric,numeric,numeric,numeric,int,int,int,int,bool,bool) as insert into ng_state_payroll_payroll_item (employee_id,active,payroll_id,gross_income,net_income,balance_income,taxable_income,paye_tax,paye_tax_annual,leave_allowance,department_id,paycategory_id,payscheme_id,level_id,retiring,resolve) values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16) returning id')
                        self.env.cr.execute('prepare insert_item_line (int, text, numeric) as insert into ng_state_payroll_payroll_item_line (item_id,name,amount) values ($1, $2, $3) returning id')
                        pay_month = datetime.strptime(self.calendar_id.from_date, '%Y-%m-%d').strftime('%m')
                        pay_year = datetime.strptime(self.calendar_id.from_date, '%Y-%m-%d').strftime('%Y')
                        
                        exception_headers = Set()
                        for emp in employees:                            
                            #_logger2.info("---------------------------------------------")
                            #_logger2.info("Name=%s,EmployeeID=%d,EmployeeNo=%s", emp.name_related, emp.id, emp.employee_no)
                                
                            #Create Payroll Item and Payroll Item Lines
                            active_flag = ('t' if emp.status_id.name == 'ACTIVE' else 'f')
                            item_line_income = 0
                            item_line_gross = 0
                            item_line_earnings_standard = 0
                            item_line_earnings_nonstd = 0
                            item_line_deductions_standard = 0
                            item_line_deductions_nonstd = 0
                            item_line_leave = 0
                            actual_leave_allowance = 0
                            item_line_deduction = 0
                            item_line_relief = 0
                            item_line_income_ded = 0
                            item_line_pension = 0
                            item_line_nhf = 0
                            item_line_other = 0
                            
                            standard_earnings = emp.standard_earnings.filtered(lambda r: r.active == True)
                            standard_deductions = emp.standard_deductions.filtered(lambda r: r.active == True)
                            nonstd_earnings = emp.nonstd_earnings.filtered(lambda r: r.active == True and ((self.calendar_id in r.calendars and r.permanent == False) or r.permanent == True))
                            nonstd_deductions = emp.nonstd_deductions.filtered(lambda r: r.active == True and ((self.calendar_id in r.calendars and r.permanent == False) or r.permanent == True))
                            loans = emp.loan_items.filtered(lambda r: r.active == True and r.state == 'approved')
                            #Calculate each standard earning
                            record_dict = {}                        
                                                    
                            basic_salary = False
                            
                            house_rent = False
                            
                            birth_month = False
                            if emp.birthday:
                                birth_month = datetime.strptime(emp.birthday, '%Y-%m-%d').strftime('%m')
                            retirement_date = False
                            item_line_retiring = 'f'
                            retirement_date_dofa = False
                            retirement_date_dob = False
                            if not emp.retirement_due_date:
                                #Pro-rate for retiring employees
                                #Use hire date and date of birth to calculate retirement date
                                if emp.payscheme_id.use_dofa:
                                    retirement_date_dofa = datetime.strptime(emp.hire_date, DEFAULT_SERVER_DATE_FORMAT) + relativedelta(years=emp.payscheme_id.service_years)
                                    retirement_date = retirement_date_dofa
                                if emp.payscheme_id.use_dob:
                                    retirement_date_dob = datetime.strptime(emp.birthday, DEFAULT_SERVER_DATE_FORMAT) + relativedelta(years=emp.payscheme_id.retirement_age)
                                    retirement_date = retirement_date_dob
                                if emp.payscheme_id.use_dofa and emp.payscheme_id.use_dob:
                                    if retirement_date_dofa < retirement_date_dob:
                                        retirement_date = retirement_date_dofa
                                    else:
                                        retirement_date = retirement_date_dob
                                if retirement_date:
                                    emp.update({'retirement_due_date':retirement_date.strftime(DEFAULT_SERVER_DATE_FORMAT)})
                                #_logger2.info("Pay Month=%s", pay_month)
                                #_logger2.info("Pay Year=%s", pay_year)
                            else:
                                retirement_date = datetime.strptime(emp.retirement_due_date, DEFAULT_SERVER_DATE_FORMAT)
    
                            if retirement_date and (int(retirement_date.month) != int(pay_month) or int(retirement_date.year) != int(pay_year)):
                                retirement_date = False
                            if retirement_date:
                                item_line_retiring = 't'
                                _logger.info("Retirement Date=%s", retirement_date)
                                #_logger2.info("Retirement Day=%s", retirement_date.day)
                                #_logger2.info("Retirement Month=%s", retirement_date.month)
                                #_logger2.info("Retirement Year=%s", retirement_date.year)
                                #_logger2.info("Retirement Date DOFA=%s", retirement_date_dofa)
                            
                            for o in standard_earnings:
                                if o.name == 'BASIC SALARY':
                                    basic_salary = o
                                if o.name == 'RENT ALLOWANCE':
                                    house_rent = o
                                amount = 0
                                if o.derived_from:
                                    amount = o.amount * o.derived_from.amount * 0.01
                                else:
                                    amount = o.amount
                                #_logger2.info("Standard Earning[%s]=%f", o.name, amount)
                                item_line_earnings_standard += amount
                                record_dict.update({o.name: (amount / 12)})
                             
                            #Calculate each standard deduction
                            for o in standard_deductions:
                                amount = 0
                                if o.derived_from:
                                    if o.derived_from.derived_from:
                                        amount = o.amount * (o.derived_from.amount * 0.01 * o.derived_from.derived_from.amount) * 0.01
                                    else:
                                        amount = o.amount * o.derived_from.amount * 0.01
                                else:
                                    amount = o.amount
                                #_logger2.info("Standard Deduction[%s]=%f", o.name, -amount)
                                item_line_deduction += amount
                                item_line_deductions_standard += amount
                                if o.income_deduction:
                                    item_line_income_ded += amount
                                    #_logger2.info("Income Ded[%s]=%f", o.name, -amount)                        
                                if o.relief and not o.income_deduction:
                                    item_line_relief += amount
                                    #_logger2.info("Relief[%s]=%f", o.name, -amount)
                                if o.name.upper().find('PENSION') >= 0:
                                    item_line_pension += amount
                                elif o.name.upper().find('NHF') >= 0:
                                    item_line_nhf += amount
                                else:
                                    item_line_other += amount
                                record_dict.update({o.name: (amount / 12)})
                                                            
                            #Calculate each non-standard earning
                            for e in nonstd_earnings:
                                item_line_earnings_nonstd += (e.amount)
                                record_dict.update({e.name: e.amount})
                                #_logger2.info("Nonstandard Earning[%s]=%f" % (e.name, e.amount))
                    
                            #Calculate each non-standard deduction
                            for d in nonstd_deductions:
                                #_logger2.info("Nonstandard Deduction[%s]=%f", d.name, -d.amount)
                                #TODO Create a configuration entity to manage reliefs
                                item_line_deduction += (d.amount * 12)
                                item_line_deductions_nonstd += (d.amount)
                                #?? Multiply by 12?
                                if d.income_deduction:
                                    item_line_income_ded += (d.amount * 12)
                                    #_logger2.info("Income Ded[%s]=%f", d.name, -d.amount)                        
                                if d.relief and not d.income_deduction:
                                    item_line_relief += (d.amount * 12)
                                    #_logger2.info("Relief[%s]=%f", d.name, -d.amount)                        
                                if d.name.upper().find('PENSION') >= 0:
                                    item_line_pension += d.amount
                                elif d.name.upper().find('NHF') >= 0:
                                    item_line_nhf += d.amount
                                else:
                                    item_line_other += d.amount
                                record_dict.update({d.name: amount})
    
                            #Calculate each loan deduction and update
                            for l in loans:
                                #_logger2.info("Loan Deduction[%s]=%f", l.name, -l.payment_amount)
                                item_line_deductions_nonstd += l.payment_amount
                                item_line_other += l.payment_amount
                                today = datetime.now().date()
                                payment_item = {
                                    'loan_id': l.id, 
                                    'date': today.strftime(DEFAULT_SERVER_DATE_FORMAT), 
                                    'amount': l.payment_amount,
                                    'employee_id': emp.id,
                                    'payroll_id': self.id,
                                }
                                if len(l.payment_ids) < l.tenure:
                                    l.write({'payment_ids': [(0, 0, payment_item)]})
                                record_dict.update({l.name: l.payment_amount})
    
                            item_id = False
                                    
                            item_line_gross = item_line_earnings_nonstd + item_line_earnings_standard
    
                            #Reduce Annual Gross by Party Deduction if Employee is Political (10% of Annual Basic)
                            item_line_gross = item_line_gross - item_line_income_ded
                                    
                            #Pay Leave Allowance for employees on birthdays that fall in this pay calendar
                            #Add Leave allowance to taxable and gross income
                            item_line_income = item_line_gross
                            leave_allowance = self.env['ng.state.payroll.leaveallowance'].search([('payscheme_id', '=', emp.payscheme_id.id)])
                            #if not leave_allowance:
                                #leave_allowance = self.env['ng.state.payroll.leaveallowance'].create({'payscheme_id':emp.payscheme_id.id,'percentage':10})
                            
                            # Capture January Grade Level and use
                            if int(pay_month) == 1:
                                emp.update({'level_id_leave_allowance': emp.level_id})
                            if not emp.level_id_leave_allowance:
                                emp.update({'level_id_leave_allowance': emp.level_id})
                            
                            # If emp.hire_date earlier than 31st July
                            hire_date = datetime.strptime(emp.hire_date, DEFAULT_SERVER_DATE_FORMAT)
                            today = date.today()
                            months_worked = relativedelta(date.today(), hire_date).years * 12 + relativedelta(date.today(), hire_date).months
                            
                            #_logger2.info("Months Worked=%d", months_worked)

                            #If at least 6 months worked
                            if emp.birthday and months_worked >= 6:
                                if leave_allowance and basic_salary:
                                    if len(leave_allowance) > 1:
                                        leave_allowance = leave_allowance.filtered(lambda r: r.paygrade_id.id == emp.level_id_leave_allowance.paygrade_id.id)
                                    if leave_allowance:
                                        if leave_allowance[0].computation_base == 'basic':
                                            item_line_leave = basic_salary.amount * leave_allowance[0].percentage / 100
                                        elif leave_allowance[0].computation_base == 'basic_rent' and house_rent:
                                            item_line_leave = (basic_salary.amount + house_rent.amount) * leave_allowance[0].percentage / 100
                                        item_line_income += item_line_leave
                                    #_logger2.info("Leave Allowance=%f", item_line_leave)
                            
                            if (emp.birthday and months_worked >= 6 and int(pay_month) == int(birth_month)) or (item_line_retiring == 't' and int(birth_month) > int(pay_month)):
                                actual_leave_allowance = item_line_leave
            
                            #Calculate PAYE Tax for each employee based on each taxable income items.
                            #TODO Use 20% of Annual Basic rather than Annual Income
                            #_logger2.info("Total Relief=%f", item_line_relief)
                            item_line_relief = item_line_relief + (item_line_income * 0.2 + 200000) #CRA relief
    
                            item_line_taxable = item_line_income - item_line_relief
                            item_line_tax = 0
                            prev_to_amount = 0
                            if item_line_taxable < 0:
                                item_line_taxable = 0
                            for taxrule in paye_taxrules:
                                if item_line_taxable - taxrule.to_amount >= 0:
                                    item_line_tax += ((taxrule.percentage / 100) * (taxrule.to_amount - prev_to_amount))
                                    #_logger2.info("Amount=%f,Percentage=%f, PAYE=%f", (taxrule.to_amount - prev_to_amount), taxrule.percentage, item_line_tax)
                                    prev_to_amount = taxrule.to_amount
                                else:
                                    item_line_tax += ((taxrule.percentage / 100) * (item_line_taxable - prev_to_amount))
                                    #_logger2.info("Amount=%f,Percentage=%f, PAYE=%f", (item_line_taxable - prev_to_amount), taxrule.percentage, item_line_tax)
                                    break
                           
                            #Apply 1% PAYE rule
                            tax_1percent = item_line_income * 0.01
                            if item_line_tax < tax_1percent:
                                item_line_tax = tax_1percent                        
                            
                            monthly_gross = 0
                            monthly_deductions = 0
                            monthly_net = 0
                            multiplication_factor = 1
                            if item_line_retiring == 't':
                                last_day = last_day_of_month(retirement_date.year, retirement_date.month)
                                day_count = retirement_date.day
                                if last_day == day_count:
                                    multiplication_factor = 1
                                else:
                                    if day_count > 30:
                                        day_count = 30
                                    multiplication_factor = float(day_count) / 30.0
                            
                            #TODO Implement T&A Rules
                            if self.apply_ta_rules:
                                #TODO Exclude employees currently on leave or who went on leave during the calendar period
                                self.env.cr.execute("select count(id) from hr_holidays where employee_id=" + str(emp.id) + " and (date_from >= '" + datetime.strptime(self.calendar_id.from_date, '%Y-%m-%d') + "' and (date_from <= '" + datetime.strptime(self.calendar_id.to_date, '%Y-%m-%d') + "' or date_to >= '" + datetime.strptime(self.calendar_id.from_date, '%Y-%m-%d') + "' and date_to <= '" + datetime.strptime(self.calendar_id.to_date, '%Y-%m-%d') + "')")
                                holiday_count = self.env.cr.fetchone()
                                self.env.cr.execute("select sum(worked_hours) from hr_attendance where employee_id=" + str(emp.id) + " and name >= '" + datetime.strptime(self.calendar_id.from_date, '%Y-%m-%d') + "' and name <= '" + datetime.strptime(self.calendar_id.to_date, '%Y-%m-%d') + "'")
                                worked_hours = self.env.cr.fetchone()
                                if holiday_count == 0:
                                    for ta_rule in ta_rules:
                                        if (worked_hours / self.calendar_id.total_hours) > ta_rule.presence:
                                            multiplication_factor *= (ta_rule.percentage / 100)
                                            break
                            
                            monthly_taxable = (item_line_taxable / 12) * multiplication_factor
                            monthly_gross = (item_line_earnings_nonstd + item_line_earnings_standard / 12) * multiplication_factor
                            monthly_deductions = item_line_deductions_nonstd + (item_line_deductions_standard / 12) * multiplication_factor
                            monthly_tax = (item_line_tax / 12) * multiplication_factor
                            monthly_net = monthly_gross - monthly_deductions - monthly_tax
                                    
                            item_line_pension = item_line_pension * multiplication_factor
                            item_line_nhf = item_line_nhf * multiplication_factor
                            item_line_other = item_line_other * multiplication_factor
                            
                            proration_prefix = ("Pro-rated " if item_line_retiring == 't' else "Full ")
                            #_logger2.info("Annual Income=%f", item_line_income)
                            #_logger2.info("Annual Gross=%f", item_line_gross)
                            #_logger2.info(proration_prefix + "Monthly Gross=%f", monthly_gross)
                            #_logger2.info("Annual Net=%f", item_line_net)
                            #_logger2.info(proration_prefix + "Monthly Net=%f", monthly_net)
                            #_logger2.info("Annual Relief=%f", item_line_relief)
                            #_logger2.info("Annual Taxable=%f", item_line_taxable)
                            #_logger2.info("Annual PAYE=%f", item_line_tax)
                            #_logger2.info(proration_prefix + "Monthly PAYE=%f", monthly_tax)
                            #_logger2.info(proration_prefix + "Monthly Deduction=%f", monthly_deductions)
                            
                            if monthly_gross > 0:
                                #TODO If monthly_net < 0 add to exception list emailed to administrator
                                monthly_net_dec = Decimal(monthly_net)
                                #monthly_net_dec.quantize(Decimal.TWOPLACES)
                                monthly_net_dec = monthly_net_dec.quantize(Decimal('.01'), rounding=ROUND_DOWN)
                                if monthly_net_dec.is_zero() or monthly_net_dec.is_signed():
                                    #Recode for Exceptional Report - add list of earnings and deductions on each row per employee
                                    record_dict.update({'name':emp.name, 'emp_no':emp.employee_no, 'dept':emp.department_id.name, 'gross':monthly_gross, 'net':monthly_net, 'taxable':(item_line_taxable / 12), 'tax':(item_line_tax / 12)})
                                    exception_list.append(record_dict)
                                    exception_headers.update(record_dict.keys())
                                
                                if emp.payscheme_id.id and emp.level_id.id:
                                    self.env.cr.execute('execute insert_item(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)', (emp.id,active_flag,self.id,monthly_gross,monthly_net,monthly_net,monthly_taxable,monthly_tax,item_line_tax,actual_leave_allowance,emp.department_id.id,emp.paycategory_id.id,emp.payscheme_id.id,emp.level_id.id,item_line_retiring,'f'))
                                    item_id = self.env.cr.fetchone()
                                else:
                                    record_dict.update({'name':emp.name, 'emp_no':emp.employee_no, 'dept':emp.department_id.name, 'gross':monthly_gross, 'net':monthly_net, 'taxable':(item_line_taxable / 12), 'tax':(item_line_tax / 12)})
                                    exception_list.append(record_dict)
                                    exception_headers.update(record_dict.keys())
                            elif emp.level_id.paygrade_id.gross_ceiling < monthly_gross:
                                record_dict.update({'name':emp.name, 'emp_no':emp.employee_no, 'dept':emp.department_id.name, 'gross':monthly_gross, 'gross_ceiling':emp.level_id.paygrade_id.gross_ceiling, 'net':monthly_net, 'taxable':(item_line_taxable / 12), 'tax':(item_line_tax / 12)})
                                exception_list.append(record_dict)
                                exception_headers.update(record_dict.keys())
                                _logger.info("Exce Head: " + repr(exception_headers))
                                _logger.info("Exception: " + repr(record_dict))
                                if emp.department_id and emp.paycategory_id and emp.payscheme_id and emp.level_id:
                                    self.env.cr.execute('execute insert_item(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)', (emp.id,active_flag,self.id,0,0,0,0,0,0,0,emp.department_id.id,emp.paycategory_id.id,emp.payscheme_id.id,emp.level_id.id,'f','t'))
                                    item_id = self.env.cr.fetchone()
                            else:
                                #TODO Recode for Exceptional Report - add list of earnings and deductions on each row per employee
                                record_dict.update({'name':emp.name, 'emp_no':emp.employee_no, 'dept':emp.department_id.name, 'gross':monthly_gross, 'net':monthly_net, 'taxable':(item_line_taxable / 12), 'tax':(item_line_tax / 12)})
                                exception_list.append(record_dict)
                                exception_headers.update(record_dict.keys())
                                _logger.info("Exce Head: " + repr(exception_headers))
                                _logger.info("Exception: " + repr(record_dict))
                                if emp.department_id and emp.paycategory_id and emp.payscheme_id and emp.level_id:
                                    self.env.cr.execute('execute insert_item(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)', (emp.id,active_flag,self.id,0,0,0,0,0,0,0,emp.department_id.id,emp.paycategory_id.id,emp.payscheme_id.id,emp.level_id.id,'f','t'))
                                    item_id = self.env.cr.fetchone()
                            if item_id:
                                #Leave allowance for the month
                                self.env.cr.execute('execute insert_item_line(%s,%s,%s)', (item_id[0],'PAYE',-monthly_tax))
                                #Leave allowance for the month
                                self.env.cr.execute('execute insert_item_line(%s,%s,%s)', (item_id[0],'Leave Allowance',actual_leave_allowance))
                                #Calculate each standard earning
                                for o in standard_earnings:
                                    amount = 0
                                    if o.derived_from:
                                        amount = o.amount * o.derived_from.amount * 0.01
                                    else:
                                        amount = o.amount
                                    self.env.cr.execute('execute insert_item_line(%s,%s,%s)', (item_id[0], o.name, (amount / 12)))
                                 
                                for o in standard_deductions:
                                    amount = 0
                                    if o.derived_from:
                                        if o.derived_from.derived_from:
                                            amount = o.amount * (o.derived_from.amount * 0.01 * o.derived_from.derived_from.amount) * 0.01
                                        else:
                                            amount = o.amount * o.derived_from.amount * 0.01
                                    else:
                                        amount = o.amount
                                    #Prorate NHF and PENSION
                                    name = o.name
                                    if retirement_date:
                                        if o.name.startswith('NHF') or o.name.startswith('PENSION'):
                                            amount = amount * multiplication_factor
                                            name = 'PRORATED ' + o.name
                                    self.env.cr.execute('execute insert_item_line(%s,%s,%s)', (item_id[0],name,(-amount / 12)))
                                        
                                for e in nonstd_earnings:
                                    self.env.cr.execute('execute insert_item_line(%s,%s,%s)', (item_id[0],'OTHER EARNINGS - ' + e.name,e.amount))
                                for e in nonstd_deductions:
                                    self.env.cr.execute('execute insert_item_line(%s,%s,%s)', (item_id[0],'OTHER DEDUCTIONS - ' + e.name,-e.amount))
                                for e in loans:
                                    self.env.cr.execute('execute insert_item_line(%s,%s,%s)', (item_id[0],'OTHER DEDUCTIONS - ' + l.name,-l.payment_amount))
                                    
                            #When an employee has been reinstated in this calendar period,- 
                            #pick all previously inactive payroll items from previous calendar- 
                            #periods from the suspension month to current calendar period, move-
                            #them to current pay period and set them active
                            reinstatement = self.env['ng.state.payroll.disciplinary'].search([('employee_id', '=', emp.id), ('action_type', '=', 'reinstatement'), ('date', '>=', self.calendar_id.from_date), ('date', '<=', self.calendar_id.to_date)])
                            
                            if reinstatement:
                                suspensions = self.env['ng.state.payroll.disciplinary'].search([('employee_id', '=', emp.id), ('action_type', '=', 'suspension')], order='date desc')
                                if len(suspensions) > 0 and not suspensions[0].unpaid_suspension:
                                    arrear_items = self.env['ng.state.payroll.payroll.item'].search([('employee_id', '=', emp.id), ('active', '=', False)])
                                    for item in arrear_items:
                                        for line_item in item.item_line_ids:
                                            line_item.write({'name':('ARREARS - ' + line_item.name + ' ' + item.payroll_id.calendar_id.name)})
                                        item.write({'active':True,'payroll_id':self.id})                                        
                                
                        self.env.cr.commit()
                        if len(exception_list) > 0:
                            with open('/tmp/payroll_exceptions_' + str(self.id) + '.csv', 'w') as csvfile:
                                writer = csv.DictWriter(csvfile, fieldnames=exception_headers)
                                writer.writeheader()
                                writer.writerows(exception_list)
                                csvfile.close()
                
                        #Process subventions
                        subvention_list = []
                        for subv in subventions:
                            subvention_list.append({'department_id': subv.org_id.id,'name': subv.name,'active': subv.active,'amount': subv.amount,'payroll_id':self.id})
                        self.write({'subvention_item_ids': [(0, 0, x) for x in subvention_list]})
    
                        _logger.info("summarize payroll count = %d", len(self.payroll_item_ids))
                        self.env.cr.execute('prepare insert_summary_payroll (int,numeric,numeric,numeric,numeric,numeric,numeric,numeric,numeric,int,int) as insert into ng_state_payroll_payroll_summary (total_strength,total_taxable_income,total_gross_income,total_net_income,total_leave_allowance,total_paye_tax,total_pension,total_nhf,total_other_deductions,department_id,payroll_id) values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) returning id')
                        self.env.cr.execute("select distinct department_id from ng_state_payroll_payroll_item where active='t' and payroll_id=" + str(self.id))
                        department_ids = self.env.cr.fetchall()
                        dept_count = 1
                        for department_id in department_ids:
                            _logger.info("Summarizing department_id = %d; %d of %d.", department_id[0], dept_count, len(department_ids))
                            self.env.cr.execute("select count(id),sum(taxable_income),sum(gross_income),sum(net_income),sum(leave_allowance),sum(paye_tax) from ng_state_payroll_payroll_item where active='t' and department_id=" + str(department_id[0]) + " and payroll_id=" + str(self.id))
                            total_strength,total_taxable_income,total_gross_income,total_net_income,total_leave_allowance,total_paye_tax = self.env.cr.fetchall()[0]
                            self.env.cr.execute("select sum(amount) from ng_state_payroll_payroll_item_line where amount < 0 and name like '%PENSION%' and item_id in (select id from ng_state_payroll_payroll_item where active='t' and department_id=" + str(department_id[0]) + " and payroll_id=" + str(self.id) + ")")
                            pension = self.env.cr.fetchone()[0]
                            if not pension:
                                pension = 0.0
                            self.env.cr.execute("select sum(amount) from ng_state_payroll_payroll_item_line where amount < 0 and name like '%NHF%' and item_id in (select id from ng_state_payroll_payroll_item where active='t' and department_id=" + str(department_id[0]) + " and payroll_id=" + str(self.id) + ")")
                            nhf = self.env.cr.fetchone()[0]
                            if not nhf:
                                nhf = 0.0
                            self.env.cr.execute("select sum(amount) from ng_state_payroll_payroll_item_line where amount < 0 and name not like '%NHF%' and name not like '%PENSION%' and item_id in (select id from ng_state_payroll_payroll_item where active='t' and department_id=" + str(department_id[0]) + " and payroll_id=" + str(self.id) + ")")
                            other = self.env.cr.fetchone()[0]
                            if not other:
                                other = 0.0
                            self.env.cr.execute('execute insert_summary_payroll(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)', (total_strength,total_taxable_income,total_gross_income,total_net_income,total_leave_allowance,total_paye_tax,pension,nhf,other,department_id[0],self.id))
                            dept_count += 1
                        self.env.cr.execute("select sum(taxable_income),sum(gross_income),sum(net_income),sum(paye_tax) from ng_state_payroll_payroll_item where active='t' and payroll_id=" + str(self.id))
                        total_taxable,total_gross,total_net,total_tax = self.env.cr.fetchone()
                        self.env.cr.execute("update ng_state_payroll_payroll set state='processed'" + ",total_net_payroll=" + str(total_net) + ",total_gross_payroll=" + str(total_gross) + ",total_tax_payroll=" + str(total_tax) + ",total_taxable_payroll=" + str(total_taxable) + ",total_balance_payroll=" + str(total_net) + ",processing_time_payroll=" + str((time.time() - tic)) + " where id=" + str(self.id))
                        self.env.cr.commit()

                        self.env.cr.execute('prepare insert_schoolsummary_payroll (int,numeric,numeric,numeric,numeric,numeric,numeric,numeric,numeric,int,int) as insert into ng_state_payroll_payroll_schoolsummary (total_strength,total_taxable_income,total_gross_income,total_net_income,total_leave_allowance,total_paye_tax,total_pension,total_nhf,total_other_deductions,school_id,payroll_id) values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) returning id')
                        self.env.cr.execute("select distinct school_id from hr_employee where id in (select distinct employee_id from ng_state_payroll_payroll_item where active='t' and payroll_id=" + str(self.id) + ")")
                        school_ids = self.env.cr.fetchall()
                        school_count = 1
                        for school_id in school_ids:
                            if school_id[0] != None:
                                _logger.info("Summarizing school_id = %d; %d of %d.", school_id[0], school_count, len(school_ids))
                                self.env.cr.execute("select org_id from ng_state_payroll_school where id=" + str(school_id[0]))
                                school_dept_ids_fetched = self.env.cr.fetchall()
                                school_dept_ids = []
                                for e in school_dept_ids_fetched:
                                    school_dept_ids.append(str(e[0]))                            
                                self.env.cr.execute("select count(id),sum(taxable_income),sum(gross_income),sum(net_income),sum(leave_allowance),sum(paye_tax) from ng_state_payroll_payroll_item where active='t' and " + str(school_id[0]) + " = (select school_id from hr_employee where id=ng_state_payroll_payroll_item.employee_id) and payroll_id=" + str(self.id))
                                total_strength,total_taxable_income,total_gross_income,total_net_income,total_leave_allowance,total_paye_tax = self.env.cr.fetchall()[0]
                                self.env.cr.execute("select sum(amount) from ng_state_payroll_payroll_item_line where amount < 0 and name like '%PENSION%' and item_id in (select id from ng_state_payroll_payroll_item where active='t' and " + str(school_id[0]) + " = (select school_id from hr_employee where id=ng_state_payroll_payroll_item.employee_id) and payroll_id=" + str(self.id) + ")")
                                pension = self.env.cr.fetchone()[0]
                                if not pension:
                                    pension = 0.0
                                self.env.cr.execute("select sum(amount) from ng_state_payroll_payroll_item_line where amount < 0 and name like '%NHF%' and item_id in (select id from ng_state_payroll_payroll_item where active='t' and " + str(school_id[0]) + " = (select school_id from hr_employee where id=ng_state_payroll_payroll_item.employee_id) and payroll_id=" + str(self.id) + ")")
                                nhf = self.env.cr.fetchone()[0]
                                if not nhf:
                                    nhf = 0.0
                                self.env.cr.execute("select sum(amount) from ng_state_payroll_payroll_item_line where amount < 0 and name not like '%NHF%' and name not like '%PENSION%' and item_id in (select id from ng_state_payroll_payroll_item where active='t' and " + str(school_id[0]) + " = (select school_id from hr_employee where id=ng_state_payroll_payroll_item.employee_id) and payroll_id=" + str(self.id) + ")")
                                other = self.env.cr.fetchone()[0]
                                if not other:
                                    other = 0.0
                                self.env.cr.execute('execute insert_schoolsummary_payroll(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)', (total_strength,total_taxable_income,total_gross_income,total_net_income,total_leave_allowance,total_paye_tax,pension,nhf,other,school_id[0],self.id))
                            school_count += 1
                        
                        if self.notify_emails:
                            _logger.info("Attempting to send emails to: " + self.notify_emails)
                            message = "Dear Sir/Madam,\nPayroll '" + self.name + "' is done preprocessing.\n\nThank you.\n"
                            message_exception = "\nPS: There were " + str(len(exception_list)) + " exceptions."
                            smtp_obj = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
                            smtp_obj.ehlo()
                            smtp_obj.login(user="osun.payroll@gmail.com", password="p@55w0rd1939")
                            
                            sender = 'osun.payroll@gmail.com'
                            receivers = self.notify_emails #Comma separated email addresses
                            msg = MIMEMultipart()
                            msg['Subject'] = 'Payroll Still in Progress' 
                            msg['From'] = sender
                            #msg['To'] = ', '.join(receivers)
                            msg['To'] = receivers
                             
                            part = False
                            if len(exception_list) > 0:
                                _logger.info("Exception count: " + str(len(exception_list)))
                                part = MIMEBase('application', "octet-stream")
                                part.set_payload(open('/tmp/payroll_exceptions_' + str(self.id) + '.csv', "rb").read())
                                Encoders.encode_base64(part)                            
                                part.add_header('Content-Disposition', 'attachment; filename="payroll_exceptions' + str(self.id) + '.csv"')
                                message = message + message_exception
                            msg.attach(MIMEText(message))
                            
                            if part:
                                _logger.info("Attaching exception report...")
                                msg.attach(part)
    
                            smtp_obj.sendmail(sender, receivers, msg.as_string())         
                            _logger.info("Email successfully sent to: " + receivers)
                        
                if self.do_pension:                    
                    tic = time.time()
                    
                    #List all pension deductions
                    deductions_pension = self.env['ng.state.payroll.deduction.pension'].search([('active', '=', True)])
                    _logger.info("Count deductions_pension= %d", len(deductions_pension))        
                    pensioners = False
                    if not self.create_user.domain_tco_types:
                        _logger.info("No Domain TCOs.")
                        pensioners = self.env['hr.employee'].search([('active', '=', True), '|', ('status_id.name', '=', 'PENSIONED'), ('status_id.name', '=', 'PENSIONED_SUSPENDED')], order='id')
                    else:
                        _logger.info("Domain TCOs= %s", self.create_user.domain_tco_types)
                        pensioner_all = self.env['hr.employee'].search([('active', '=', True), '|', ('status_id.name', '=', 'PENSIONED'), ('status_id.name', '=', 'PENSIONED_SUSPENDED')], order='id')
                        pensioners = []
                        for ps in pensioner_all:
                            for tco_type in self.create_user.domain_tco_types:
                                if tco_type.active and tco_type.tco_id.id == ps.tco_id.id and tco_type.pensiontype_id.id == ps.pensiontype_id.id:
                                    pensioners.append(ps)
                                    break
        
                    _logger.info("Count pensioners= %d", len(pensioners))        
                    
                    exception_list = []
                    self.env.cr.execute('prepare insert_item_pension (int,bool,int,int,numeric,numeric,numeric,numeric) as insert into ng_state_payroll_pension_item (employee_id,active,payroll_id,tco_id,gross_income,net_income,balance_income,arrears_amount) values ($1,$2,$3,$4,$5,$6,$7,$8) returning id')
                    self.env.cr.execute('prepare insert_item_line_pension (int, text, numeric) as insert into ng_state_payroll_pension_item_line (item_id,name,amount) values ($1,$2,$3) returning id')
                    for pen in pensioners:
                        pension_amount = pen.annual_pension / 12
                        gross_amount = 0
                        gross_amount = gross_amount + pension_amount
                        earnings_nonstd = 0
                        deductions_nonstd = 0
                        active_flag = ('t' if pen.status_id.name == 'PENSIONED' else 'f')
                        arrears = pen.pension_arrears.filtered(lambda r: r.active == True and self.calendar_id.id == r.calendar_id.id)
                        for ern in arrears:
                            earnings_nonstd = earnings_nonstd + ern.amount

                        
                        #When a pensioner has been reinstated in this calendar period,- 
                        #pick all previously inactive payroll items from previous calendar- 
                        #periods from the suspension month to current calendar period, move-
                        #them to current pay period and set them active
                        reinstatement = self.env['ng.state.payroll.disciplinary'].search([('employee_id', '=', pen.id), ('action_type', '=', 'reinstatement'), ('date', '>=', self.calendar_id.from_date), ('date', '<=', self.calendar_id.to_date)])
                        
                        if reinstatement:
                            suspensions = self.env['ng.state.payroll.disciplinary'].search([('employee_id', '=', pen.id), ('action_type', '=', 'suspension')], order='date desc')
                            if len(suspensions) > 0 and not suspensions[0].unpaid_suspension:
                                arrear_items = self.env['ng.state.payroll.pension.item'].search([('employee_id', '=', pen.id), ('active', '=', False)])
                                for item in arrear_items:
                                    item.write({'active':True,'payroll_id':self.id})
                                    for line_item in item.item_line_ids:
                                        if line_item.amount >= 0:
                                            earnings_nonstd = earnings_nonstd + line_item.amount
                                        else:
                                            deductions_nonstd = deductions_nonstd - line_item.amount
                                        line_item.write({'name':('ARREARS - ' + line_item.name + ' ' + item.payroll_id.calendar_id.name)})

                        gross_amount = gross_amount + earnings_nonstd

                        for ded in deductions_pension:
                            ded_amount = 0
                            if len(ded.whitelist_ids) > 0:
                                if pen.id in ded.whitelist_ids.ids:
                                    if not ded.fixed:
                                        ded_amount = gross_amount * ded.amount / 100
                                    else:
                                        ded_amount = ded.amount
                            else:
                                if not pen.id in ded.blacklist_ids.ids:
                                    if not ded.fixed:
                                        ded_amount = gross_amount * ded.amount / 100
                                    else:
                                        ded_amount = ded.amount
                            deductions_nonstd = deductions_nonstd + ded_amount
                        if pension_amount < 0:
                            exception_list.append({'name':pen.name, 'gross':gross_amount, 'net':pension_amount})
                        net_amount = gross_amount - deductions_nonstd
                        self.env.cr.execute('execute insert_item_pension(%s,%s,%s,%s,%s,%s,%s,%s)', (pen.id,active_flag,self.id,pen.tco_id.id,gross_amount,net_amount,net_amount,earnings_nonstd))
                        item_id = self.env.cr.fetchone()
                        self.env.cr.execute('execute insert_item_line_pension(%s,%s,%s)', (item_id[0],'Monthly Pension',pension_amount))

                        for ern in arrears:
                            self.env.cr.execute('execute insert_item_line_pension(%s,%s,%s)', (item_id[0],'ARREARS - ' + ern.name,ern.amount))

                        for ded in deductions_pension:
                            ded_amount = 0
                            if len(ded.whitelist_ids) > 0:
                                if pen.id in ded.whitelist_ids.ids:
                                    if not ded.fixed:
                                        ded_amount = gross_amount * ded.amount / 100
                                    else:
                                        ded_amount = ded.amount
                                    #_logger.info("WhiteList Deduction (%f) '%s'=%f %f", gross_amount, ded.name, ded.amount, ded_amount)
                                    if ded_amount > 0:
                                        self.env.cr.execute('execute insert_item_line_pension(%s,%s,%s)', (item_id[0],ded.name,-ded_amount))
                            else:
                                if not pen.id in ded.blacklist_ids.ids:
                                    if not ded.fixed:
                                        ded_amount = gross_amount * ded.amount / 100
                                    else:
                                        ded_amount = ded.amount
                                    #_logger.info("BlackList Deduction (%f) '%s'=%f %f", gross_amount, ded.name, ded.amount, ded_amount)
                                    if ded_amount > 0:
                                        self.env.cr.execute('execute insert_item_line_pension(%s,%s,%s)', (item_id[0],ded.name,-ded_amount))

                    self.env.cr.commit()
                    
                    if len(exception_list) > 0:
                        with open('/tmp/pension_exceptions.csv', 'w') as csvfile:
                            fieldnames = ['name', 'gross', 'net', 'taxable', 'tax']
                            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                            writer.writeheader()
                            writer.writerows(exception_list)
                            csvfile.close()
                        
                    _logger.info("summarize pension count = %d", len(self.pension_item_ids))
                    self.env.cr.execute('prepare insert_summary_pension (int,numeric,numeric,numeric,numeric,int,int) as insert into ng_state_payroll_pension_summary (total_strength,total_gross_income,total_net_income,total_arrears,total_dues,tco_id,payroll_id) values ($1,$2,$3,$4,$5,$6,$7) returning id')
                    self.env.cr.execute("select distinct tco_id from ng_state_payroll_pension_item where active='t' and payroll_id=" + str(self.id))
                    tco_ids = self.env.cr.fetchall()
                    tco_count = 1
                    for tco_id in tco_ids:
                        _logger.info("Summarizing tco_id = %d; %d of %d.", tco_id[0], tco_count, len(tco_ids))
                        self.env.cr.execute("select count(id),sum(gross_income),sum(net_income) from ng_state_payroll_pension_item where active='t' and tco_id=" + str(tco_id[0]) + " and payroll_id=" + str(self.id))
                        total_strength,total_gross_income,total_net_income = self.env.cr.fetchone()
                        self.env.cr.execute("select sum(amount) from ng_state_payroll_pension_item_line where name like '%ARREARS%' and item_id in (select id from ng_state_payroll_pension_item where active='t' and tco_id=" + str(tco_id[0]) + " and payroll_id=" + str(self.id) + ")")
                        arrears = self.env.cr.fetchone()[0]
                        if not arrears:
                            arrears = 0.0
                        self.env.cr.execute("select sum(amount) from ng_state_payroll_pension_item_line where name like '%NUP%' and item_id in (select id from ng_state_payroll_pension_item where active='t' and tco_id=" + str(tco_id[0]) + " and payroll_id=" + str(self.id) + ")")
                        dues = self.env.cr.fetchone()[0]
                        if not dues:
                            dues = 0.0
                        self.env.cr.execute('execute insert_summary_pension(%s,%s,%s,%s,%s,%s,%s)', (total_strength,total_gross_income,total_net_income,arrears,dues,tco_id[0],self.id))
                        tco_count += 1
                    self.env.cr.execute("select sum(gross_income),sum(net_income) from ng_state_payroll_pension_item where active='t' and payroll_id=" + str(self.id))
                    total_gross,total_net = self.env.cr.fetchone()
                    self.env.cr.execute("update ng_state_payroll_payroll set state='processed'" + ",total_net_pension=" + str(total_net) + ",total_gross_pension=" + str(total_gross) + ",total_balance_pension=" + str(total_net) + ",processing_time_pension=" + str((time.time() - tic)) + " where id=" + str(self.id))
                    self.env.cr.commit()
                    
                    if self.notify_emails:
                        message = "Dear Sir/Madam,\nPayroll '" + self.name + "' has completed.\n\nThank you.\n"
                        message_exception = "\nPS: There were " + str(len(exception_list)) + " exceptions.\n"
                        smtp_obj = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
                        smtp_obj.ehlo()
                        smtp_obj.login(user="osun.payroll@gmail.com", password="p@55w0rd1939")
                        
                        sender = 'osun.payroll@gmail.com'
                        receivers = self.notify_emails #Comma separated email addresses
                        msg = MIMEMultipart()
                        msg['Subject'] = 'Payroll Still in Progress' 
                        msg['From'] = sender
                        #msg['To'] = ', '.join(receivers)
                        msg['To'] = receivers
                        
                        part = False
                        if len(exception_list) > 0:
                            part = MIMEBase('application', "octet-stream")
                            part.set_payload(open("/tmp/pension_exceptions.csv", "rb").read())
                            Encoders.encode_base64(part)                            
                            part.add_header('Content-Disposition', 'attachment; filename="payroll_exceptions.csv"')
                            message = message + message_exception
                        msg.attach(MIMEText(message))
                        
                        if part:
                            msg.attach(part)
                                                
                        smtp_obj.sendmail(sender, receivers, msg.as_string())         
                        _logger.info("Email successfully sent to: " + receivers)

                    
        return True    
                            
    @api.multi
    def finalize2(self, context=None):
        _logger.info("Calling finalize2...state = %s", self.state)
        
        #TODO Change ORM create and update calls to DB Cursor SQL calls
        
        if self.state == 'in_progress':
            if self.calendar_id:
                if self.do_payroll:
                    tic = time.time()
                    item_list = []
            
                    #List all tax rules
                    paye_taxrules = self.env['ng.state.payroll.taxrule'].search([('active', '=', 'True')])
                    
                    #Fetch all active employees *TODO* (and non-suspended employees)
                    employees = self.env['hr.employee'].search([('active', '=', 'True'), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id')
                    #employees = self.env['hr.employee'].search([('employee_no', 'in', ['65754','21978','37449','20925','11355','4599','22560','17771','4414','44552','30801','35323','32546','46842','31692','14867','58081','38892','10486','43932','40052','41927','32342','12507','7037','33280','41948','41944','17966','33498','32850','41969','79972','44957','16617','41520','36962','40917','47375','39139','41954','75063','32117','44436','41641','35637','44032','27535','78315','89000','73796','72297','88006','33686','31971','72773','72759','17781','17006','47208','65679','47929','87878','5107','5113','5084','05147_D','10868','12463','74083','84370','35370','6122','6123','6119','6148','1980','1936','1957','12460','33746','32017','75184','PP199','12026','12461','89040','86900','86904','86899','86897','86887','86889','86905','86902','86903','78632','34635','34700','34581','33514','34580','10616','4196','4153','4173','78404','73200','72110','72083','12090','40435','41637','3659','89014','5102','4146','4115','4144','4184','4075','4160','38250','44058','40933','10573','6449','5019','44071','43900','43905','43904','3654','3662','32454','33770','34609','31901','36897','34512','34479','31578','36384','34608','34489','30538','34835','38225','41797','88916','88893','86268','70972','44930','44055','PP196','PP197','PP198','78219','87195','87121','88751','88862','87056','88752','78298','87028','87196','87200','87217','88806','88396','78105','78167','88791','78265','86574','86591','88567','88662','88622','88568','78322','78324','78323','87036','87046','84001','88812','88813','78120','78119','78337','78209','88698','86461','86623','86621','86847','88378','78218','86525','86522','88366','86592','5099','5103','5078','74728','4182','41087','42024','72001','86028','36309','35622','86160','78217','87119','87183','87110','87142','88740','87067','88861','87097','88746','88738','87199','87118','87136','87108','87104','87211','87141','88674','88402','88403','88643','88520','88641','88486','88497','86520','86575','88481','86521','86506','88552','86515','88526','88387','88386','86236','87209','87206','78196','78184','87155','78197','78200','87022','78125','87124','88869','88796','87205','78124','87203','88736','87191','78310','78145','78144','88795','78134','88808','88850','78140','78135','78147','78148','78306','78253','78250','78252','78266','88734','88877','78171','87077','87081','78237','78228','88864','78172','78256','78247','87083','87087','87088','78244','87082','78233','88711','88684','86576','88710','86517','88547','88314','88658','88487','86518','88329','88652','86465','88599','86485','86486','88483','88610','88612','86484','88680','86494','86467','86512','86593','86012','78364','86251','86011','88917','88991','86242','86237','86267','88008','78389','88499','88642','88644','86587','86483','88544','88562','88659','88929','86273','86293','87030','78081','87149','87150','87147','78082','78121','78117','87134','78080','88857','87094','87044','87033','87055','87042','78112','78106','88866','78113','78111','88786','78341','87076','78343','88510','88664','88651','88450','88608','88470','88511','86510','86513','88513','86490','86504','88522','88976','86255','86247','88985','86294','86252','86222','86239','86269','88007','87176','86853','86852','78204','88872','88873','88797','87140','78215','78213','87175','87174','86250','86135','86285','86202','86171','86254','78386','86095','86162','86323','88712','88706','88654','88488','88576','88293','88705','88704','88655','88975','86168','86078','86169','86185','86024','86052','78414','86138','86136','86305','86300','87169','87168','86843','86842','86275','86627','86620','86141','86089','86292','86302','86638','86612','86622','86856','86860','86846','86316','86473','86629','86284','89051','89269','87171','87166','86240','86819','88673','89267','86848','88479','86260','86023','86201','88003','78246','86161','86244','78388','88896','88894','86155','78206','88731','78153','78377','88753','88019','88750','86079','86034','88913','86125','86087','88930','86199','88895','86499','78293','88732','86180','86597','86468','86523','88582','88860','86179','86139','86073','86235','86033','86154','86231','86232','88020','86322','86845','78395','86590','86464','88766','86571','86498','86812','86818','78270','86480','87072','87074','40043','67666','67670','72129','79905','67747','71814','72133','70975','66990','38218','38216','38219','42285','5100','5130','5068','5083','5098','5039','10598','10599','4150','40151','39322','89314','84071','84072','84564','84062','89860','89865','71444','71484','71486','88210','71482','71479','71446','71447','71445','71483','71485','71481','71480','71487','5049','75159','87557','78745','79954','70254','66996','66992','87564','67751','67669','71815','72141','71811','67654','72089','72021','68494_D','71998','75157','78757','44813','38959','39170','74692','74687','70580','70590','70591','74839','74701','68387','74696','78663','78664','75026','72086','45697','45688','65689','58514','41514','36320','30994','74675','44786','72816','41655','44639','58065','40542','38913','42538','27711','38455','25144','38560','41916','41363','7871','38660','121','6340','43641','42859','42936','6289','6299','6344'])])

                    _logger.info("Count employees= %d", len(employees))        
                    
                    total_gross = 0
                    total_net = 0
                    total_tax = 0
                    total_taxable = 0
                    
                    dept_summary = {}
            
                    for emp in employees:
                        #TODO When an employee has been reinstated in this calendar period,- 
                        #pick all previously inactive payroll items from previous calendar- 
                        #periods from the suspension month to current calendar period, move-
                        #them to current pay period and set them active
                        #_logger2.info("---------------------------------------------")
                        #_logger2.info("Name=%s", emp.name_related)
                        item_list_lines = []
                            
                        #Create Payroll Item and Payroll Item Lines
                        active_flag = False
                        if emp.status_id.name == 'ACTIVE':
                            active_flag = True
                        item_dict = {'employee_id':emp.id, 'active': active_flag, 'payroll_id': self.id}
                        item_line_gross = 0
                        item_line_leave = 0
                        item_line_deduction = 0
                        item_line_taxable = 0
                        item_line_relief = 200000
                        
                        #Calculate each standard earning
                        #for o in emp.standard_earnings.filtered(lambda r: r.active == True):
                        for o in emp.employee_earnings.filtered(lambda r: r.active == True):
                            amount = 0
                            if o.fixed:
                                amount = o.amount
                            else:
                                amount = o.amount * o.derived_from.amount * 0.01
                            item_line_gross += amount
                            #_logger2.info("Standard Earning[%s]=%f", o.name, amount)
                            item_list_lines.append({'name':o.name, 'amount':(amount / 12)})
                         
                        #Calculate each standard deduction
                        #for o in emp.standard_deductions.filtered(lambda r: r.active == True):
                        for o in emp.employee_deductions.filtered(lambda r: r.active == True):
                            amount = 0
                            if o.fixed:
                                amount = -o.amount
                            else:
                                if o.derived_from.fixed:
                                    amount = o.amount * o.derived_from.amount * 0.01
                                else:
                                    amount = o.amount * (o.derived_from.amount * 0.01 * o.derived_from.derived_from.amount) * 0.01
                            #_logger2.info("Standard Deduction[%s]=%f", o.name, amount)
                            item_line_deduction += -amount
                            item_list_lines.append({'name':o.name, 'amount':(amount / 12)})

                        #_logger2.info("Gross Income=%f", item_line_gross)
                        
                        percent_1percent = item_line_gross * 0.01
                        if percent_1percent > 200000:
                            item_line_relief = percent_1percent
                        item_line_relief += (item_line_gross * 0.2) #CRA relief
                        item_line_taxable = item_line_gross - item_line_relief
                        if item_line_taxable < 0:
                            item_line_taxable = 0
                        #Calculate PAYE Tax for each employee based on each taxable income items.
                        #TODO item_line_deduction should actually be reliefs - NHF, Pension, Party
                        total_taxable += (item_line_taxable / 12)
                        item_line_tax = 0
                        if item_line_taxable > 0:
                            prev_to_amount = 0
                            for taxrule in paye_taxrules:
                                if item_line_taxable - taxrule.to_amount >= 0:
                                    item_line_tax += ((taxrule.percentage / 100) * (taxrule.to_amount - prev_to_amount))
                                    #_logger2.info("Amount=%f,Percentage=%f, PAYE=%f", (taxrule.to_amount - prev_to_amount), taxrule.percentage, item_line_tax)
                                    prev_to_amount = taxrule.to_amount
                                else:
                                    item_line_tax += ((taxrule.percentage / 100) * (item_line_taxable - prev_to_amount))
                                    #_logger2.info("Amount=%f,Percentage=%f, PAYE=%f", (item_line_taxable - prev_to_amount), taxrule.percentage, item_line_tax)
                                    break
                           
                            #Apply 1% PAYE rule
                            tax_1percent = item_line_gross * 0.01
                            if item_line_tax < tax_1percent:
                                item_line_tax = tax_1percent                        
                        
                        item_line_net = item_line_gross - item_line_deduction - item_line_tax
                        
                        #_logger2.info("Gross=%f", item_line_gross)
                        #_logger2.info("Net=%f", item_line_net)
                        #_logger2.info("Taxable=%f", item_line_taxable)
                        #_logger2.info("PAYE=%f", item_line_tax)
                        
                        item_dict.update({'item_line_ids':item_list_lines})
                        if item_line_gross > 0:
                            total_gross = total_gross + (item_line_gross / 12)
                            total_net = total_net + (item_line_net / 12)
                            total_tax = total_tax + (item_line_tax / 12)
            
                            item_dict.update({'gross_income':(item_line_gross / 12),'net_income':(item_line_net / 12),'balance_income':(item_line_net / 12),'taxable_income':(item_line_taxable / 12),'paye_tax':(item_line_tax / 12),'leave_allowance':(item_line_leave / 12),'active':emp.active})
            
                            if not dept_summary.has_key(emp.department_id.id):
                                dept_summary[emp.department_id.id] = {'department_id':emp.department_id.id,'payroll_id':self.id,'total_taxable_income':0,'total_gross_income':0,'total_net_income':0,'total_paye_tax':0,'total_leave_allowance':0}

                            dept_summary[emp.department_id.id]['total_taxable_income'] += (item_line_taxable / 12)
                            dept_summary[emp.department_id.id]['total_gross_income'] += (item_line_gross / 12)
                            dept_summary[emp.department_id.id]['total_net_income'] += (item_line_net / 12)
                            dept_summary[emp.department_id.id]['total_paye_tax'] += (item_line_tax / 12)
                            dept_summary[emp.department_id.id]['total_leave_allowance'] += (item_line_leave / 12)
                        else:
                            item_dict.update({'gross_income':(item_line_gross / 12),'net_income':(item_line_net / 12),'balance_income':(item_line_net / 12),'taxable_income':(item_line_taxable / 12),'paye_tax':(item_line_tax / 12),'leave_allowance':(item_line_leave / 12),'active':emp.active,'resolve':True})
                            
                        item_list.append(item_dict)
                    
                    self.update({'payroll_item_ids':item_list})
                    self.update({'state':'processed','total_net_payroll':total_net,'total_balance_payroll':total_net,'total_gross_payroll':total_gross,'total_tax_payroll':total_tax,'total_taxable_payroll':total_taxable, 'processing_time_payroll':(time.time() - tic)})
                    self.update({'payroll_summary_ids':dept_summary.values()})
                    
                if self.do_pension:                    
                    tic = time.time()
                    total_gross = 0
                    total_net = 0
                    
                    #List all pension deductions
                    deductions_pension = self.env['ng.state.payroll.deduction.pension'].search([('active', '=', 'True')])
                    _logger.info("Count deductions_pension= %d", len(deductions_pension))        
        
                    pensioners = self.env['hr.employee'].search([('active', '=', 'True'), ('status_id.name', '=', 'PENSIONED')])
                    _logger.info("Count pensioners= %d", len(pensioners))        
                    
                    item_list = []
    
                    for pen in pensioners:
                        item_list_lines = []                
                        pension_amount = pen.annual_pension / 12
                        gross_amount = pension_amount
                        active_flag = False
                        if pen.status_id.name == 'PENSIONED':
                            active_flag = True
                        item_dict = {'employee_id':pen.id, 'active':active_flag, 'payroll_id':self.id, 'gross_income':pension_amount}
                        ded_amount = 0
                        item_list_lines.append({'name':'Monthly Pension', 'amount':pension_amount})
                        for ded in deductions_pension:
                            if len(ded.whitelist_ids) > 0 and pen in ded.whitelist_ids:
                                if not ded.fixed:
                                    ded_amount = pension_amount * ded.amount / 100
                                else:
                                    ded_amount = ded.amount
                            if not pen in ded.blacklist_ids:
                                if not ded.fixed:
                                    ded_amount = pension_amount * ded.amount / 100
                                else:
                                    ded_amount = ded.amount
                            pension_amount -= ded_amount
                            item_list_lines.append({'name':ded.name, 'amount':-ded_amount})
                        total_gross = total_gross + gross_amount
                        total_net = total_net + pension_amount                            
                        item_dict.update({'item_line_ids':item_list_lines, 'gross_income':gross_amount, 'net_income':pension_amount, 'balance_income':pension_amount})
                        item_list.append(item_dict)
                    
                    self.update({'pension_item_ids':item_list})                                    
                    self.update({'state':'processed','total_net_pension':total_net,'total_balance_pension':total_net,'total_gross_pension':total_gross, 'processing_time_pension':(time.time() - tic)})

                    #Send email to caller on completion
                    message = ("Dear Sir/Madam,\nPayroll '" + self.name + "' has completed. Thank you.\n")
                    vals = {'state': 'outgoing',
                            'subject': 'Payroll Processing Completed',
                            'body_html': '<pre>%s</pre>' % message,
                            'email_to': 'neeyeed@gmail.com',
                            #'email_to': self.env.user.email,
                            'email_from': 'neeyeed@gmail.com',
                    }
                    email_id = self.env['mail.mail'].create(vals, context=context)
                    self.env['mail.mail'].send(email_id)
        return True                    
            
    def dry_run(self):
        _logger.info("Calling finalize...state = %s", self.state)
        if self.in_progress:
            raise osv.except_osv(_('Info'), _('Processing already in progress.'))
                    
        prev_state = self.state
        if not self.state == 'in_progress':        
            if self.calendar_id and self.state == 'draft':
                self.set_in_progress()
                
                if self.do_payroll:
                    tic = time.time()
                    item_list = []
                            
                    #List all subvention earnings for this calendar period
                    subventions = self.env['ng.state.payroll.subvention'].search([('active', '=', 'True'), ('calendar_id', '=', self.calendar_id.id)])
            
                    #List all tax rules
                    paye_taxrules = self.env['ng.state.payroll.taxrule'].search([('active', '=', 'True')])
                    
                    #List all standard earnings for this calendar period
                    earnings_nonstd_all = self.env['ng.state.payroll.earning.nonstd'].search([('active', '=', 'True'), ('calendars.id', '=', self.calendar_id.id)], order='employee_id')
                    _logger.info("Count earnings_nonstd= %d", len(earnings_nonstd_all))        
                    
                    #List all non-standard deductions for this calendar period
                    deductions_nonstd_all = self.env['ng.state.payroll.deduction.nonstd'].search([('active', '=', 'True'), ('calendars.id', '=', self.calendar_id.id)], order='employee_id')
                    _logger.info("Count deductions_nonstd= %d", len(deductions_nonstd_all))        
                    
                    #Fetch all active employees *TODO* (and non-suspended employees)
                    employees = self.env['hr.employee'].search([('active', '=', 'True'), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id')
                    _logger.info("Count employees= %d", len(employees))        
                    
                    total_gross = 0
                    total_net = 0
                    total_tax = 0
                    total_taxable = 0
                    
                    dept_summary = {}
            
                    for emp in employees:
                        #TODO When an employee has been reinstated in this calendar period,- 
                        #pick all previously inactive payroll items from previous calendar- 
                        #periods from the suspension month to current calendar period, move-
                        #them to current pay period and set them active
                        item_list_lines = []
                
                        earnings_standard = self.env['ng.state.payroll.earning.standard'].search([('active', '=', 'True'), ('payscheme_id', '=', emp.payscheme_id.id), ('level_id', '=', emp.level_id.id)])
                        deductions_standard = self.env['ng.state.payroll.deduction.standard'].search([('active', '=', 'True'), ('payscheme_id', '=', emp.payscheme_id.id), ('level_id', '=', emp.level_id.id)])
                            
                        #Create Payroll Item and Payroll Item Lines
                        active_flag = False
                        if emp.status_id.name == 'ACTIVE':
                            active_flag = True
                        item_dict = {'employee_id':emp.id, 'active': active_flag, 'payroll_id': self.id}
                        item_line_income = 0
                        item_line_gross = 0
                        item_line_leave = 0
                        item_line_deduction = 0
                        item_line_relief = 0
                        item_line_earnings_standard = 0
                        item_line_earnings_nonstd = 0
                        item_line_deductions_standard = 0
                        item_line_deductions_nonstd = 0
                        
                        basic_salary = False
                        
                        retirement_date = False
                        item_line_retiring = 'f'
                        retirement_date_dofa = False
                        retirement_date_dob = False
                        birth_month = datetime.strptime(emp.birthday, '%Y-%m-%d').strftime('%m')
                        pay_month = datetime.strptime(self.calendar_id.from_date, '%Y-%m-%d').strftime('%m')
                        pay_year = datetime.strptime(self.calendar_id.from_date, '%Y-%m-%d').strftime('%Y')
                        
                        if not emp.retirement_due_date:
                            #Pro-rate for retiring employees
                            #Use hire date and date of birth to calculate retirement date
                            if emp.payscheme_id.use_dofa:
                                retirement_date_dofa = datetime.strptime(emp.hire_date, DEFAULT_SERVER_DATE_FORMAT) + relativedelta(years=emp.payscheme_id.service_years)
                                retirement_date = retirement_date_dofa
                            if emp.payscheme_id.use_dob:
                                retirement_date_dob = datetime.strptime(emp.birthday, DEFAULT_SERVER_DATE_FORMAT) + relativedelta(years=emp.payscheme_id.retirement_age)
                                retirement_date = retirement_date_dob
                            if emp.payscheme_id.use_dofa and emp.payscheme_id.use_dob:
                                if retirement_date_dofa < retirement_date_dob:
                                    retirement_date = retirement_date_dofa
                                else:
                                    retirement_date = retirement_date_dob
                            if retirement_date:
                                emp.update({'retirement_due_date':retirement_date.strftime(DEFAULT_SERVER_DATE_FORMAT)})
                            #_logger2.info("Pay Month=%s", pay_month)
                            #_logger2.info("Pay Year=%s", pay_year)
                        else:
                            retirement_date = datetime.strptime(emp.retirement_due_date, DEFAULT_SERVER_DATE_FORMAT)

                        if retirement_date and (int(retirement_date.month) != int(pay_month) or int(retirement_date.year) != int(pay_year)):
                            retirement_date = False
                        
                        #Calculate each standard earning
                        for o in earnings_standard:
                            if o.name == 'BASIC SALARY':
                                basic_salary = o
                            amount = 0
                            if o.fixed:
                                amount = o.amount
                            else:
                                amount = o.amount * o.derived_from.amount * 0.01
                            ##_logger2.info("Standard Earning[%s]=%f", o.name, amount)
                            item_list_lines.append({'name':o.name, 'amount':amount})
                            item_line_gross += amount
                            item_line_earnings_standard += amount
                            #if o.taxable:
                                #item_line_taxable += amount
                         
                        #Calculate each standard deduction
                        for o in deductions_standard:
                            amount = 0
                            if o.fixed:
                                amount = o.amount
                            else:
                                if o.derived_from.fixed:
                                    amount = o.amount * o.derived_from.amount * 0.01
                                else:
                                    amount = o.amount * (o.derived_from.amount * 0.01 * o.derived_from.derived_from.amount) * 0.01
                            ##_logger2.info("Standard Deduction[%s]=%f", o.name, -amount)
                            item_list_lines.append({'name':o.name, 'amount':-amount})
                            item_line_deduction += amount
                            item_line_deductions_standard += amount
                            if o.name.startswith('PENSION FROM') or o.name == 'NHF' or o.name == 'PARTY DEDUCTION':
                                item_line_relief += amount
                                ##_logger2.info("Relief[%s]=%f", o.name, amount)
                                  
                            #item_line_taxable -= amount 
                                
                        earnings_nonstd = earnings_nonstd_all.filtered(lambda r: r.employee_id.id == emp.id)
                        earnings_nonstd_all = earnings_nonstd_all - earnings_nonstd
                        #Calculate each non-standard earning
                        for e in earnings_nonstd:
                            item_line_gross += (e.amount)
                            item_line_earnings_nonstd += (e.amount)
                            #if earnings_nonstd[idx_nonstd_earnings].taxable:
                                #item_line_taxable += earnings_nonstd[idx_nonstd_earnings].amount
                            ##_logger2.info("Nonstandard Earning[%s]=%f", e.name, (e.amount))
                            item_list_lines.append({'name':e.name, 'amount':(e.amount)})
                
                        deductions_nonstd = deductions_nonstd_all.filtered(lambda r: r.employee_id.id == emp.id)
                        deductions_nonstd_all = deductions_nonstd_all - deductions_nonstd
                        #Calculate each non-standard deduction
                        for d in deductions_nonstd:
                            ##_logger2.info("Nonstandard Deduction[%s]=%f", d.name, (d.amount))
                            #item_line_taxable -= deductions_nonstd[idx_nonstd_deductions].amount
                            #TODO Create a configuration entity to manage reliefs
                            item_line_deduction += (d.amount)
                            item_line_deductions_nonstd += (d.amount)
                            if d.name.startswith('PENSION FROM') or d.name == 'NHF' or d.name == 'PARTY DEDUCTION':
                                item_line_relief += (d.amount)
                                ##_logger2.info("Relief[%s]=%f", d.name, (d.amount))
                                
                        #Pay Leave Allowance for employees on birthdays that fall in this pay calendar
                        #Add Leave allowance to taxable and gross income
                        item_line_leave = 0
                        leave_allowance = self.env['ng.state.payroll.leaveallowance'].search([('payscheme_id', '=', emp.payscheme_id.id)])
                        if leave_allowance and basic_salary:
                            item_line_leave = basic_salary.amount * leave_allowance.percentage / 100
                            item_line_income += (item_line_leave + item_line_gross)
                            ##_logger2.info("Leave Allowance=%f", item_line_leave)
        
                        ##_logger2.info("Annual Income=%f", item_line_income)
                                
                        #Pro-rate for retiring employees
                        item_line_retiring = False
                        ##_logger2.info("Retirement Date=%s", emp.retirement_due_date)

                        multiplication_factor = 1
                        if item_line_retiring == 't':
                            last_day = last_day_of_month(retirement_date.year, retirement_date.month)
                            day_count = retirement_date.day
                            if last_day == day_count:
                                multiplication_factor = 1
                            else:
                                if day_count > 30:
                                    day_count = 30
                                multiplication_factor = float(day_count) / 30.0
                        
                        if emp.retirement_due_date:
                            retirement_month = datetime.strptime(emp.retirement_due_date, '%Y-%m-%d').strftime('%m')
                            retirement_year = datetime.strptime(emp.retirement_due_date, '%Y-%m-%d').strftime('%Y')
                            pay_month = datetime.strptime(self.calendar_id.from_date, '%Y-%m-%d').strftime('%m')
                            pay_year = datetime.strptime(self.calendar_id.from_date, '%Y-%m-%d').strftime('%Y')
                            if retirement_month == pay_month and retirement_year == pay_year:
                                item_line_retiring = True
                                retirement_day = datetime.strptime(emp.retirement_due_date, '%Y-%m-%d').strftime('%d')
                                item_line_gross *= multiplication_factor
                                ##_logger2.info("Pro-rated Gross=%f", item_line_gross)
                                #item_line_taxable *= (float(retirement_day) / 30)
                            
                        #Calculate PAYE Tax for each employee based on each taxable income items.
                        #TODO item_line_deduction should actually be reliefs - NHF, Pension, Party
                        item_line_relief += (item_line_income * 0.2 + 200000) #CRA relief
                        item_line_taxable = item_line_income - item_line_relief
                        total_taxable += (item_line_taxable / 12)
                        item_line_tax = 0
                        prev_to_amount = 0
                        for taxrule in paye_taxrules:
                            if item_line_taxable - taxrule.to_amount >= 0:
                                item_line_tax += ((taxrule.percentage / 100) * (taxrule.to_amount - prev_to_amount))
                                ##_logger2.info("Amount=%f,Percentage=%f, PAYE=%f", (taxrule.to_amount - prev_to_amount), taxrule.percentage, item_line_tax)
                                prev_to_amount = taxrule.to_amount
                            else:
                                item_line_tax += ((taxrule.percentage / 100) * (item_line_taxable - prev_to_amount))
                                ##_logger2.info("Amount=%f,Percentage=%f, PAYE=%f", (item_line_taxable - prev_to_amount), taxrule.percentage, item_line_tax)
                                break
                        
                        #Apply 1% PAYE rule
                        tax_1percent = item_line_income * 0.01
                        if item_line_tax < tax_1percent:
                            item_line_tax = tax_1percent 
                        
                        item_line_net = item_line_gross - item_line_deduction - item_line_tax
                        
                        monthly_gross = 0
                        monthly_deductions = 0
                        monthly_net = 0
                        if item_line_retiring:
                            day_count = retirement_date.day
                            if day_count > 30:
                                day_count = 30
                            monthly_gross = (item_line_earnings_nonstd + item_line_earnings_standard / 12) * float(day_count) / 30.0
                            monthly_deductions = (item_line_deductions_nonstd + item_line_deductions_standard / 12) * float(day_count) / 30.0
                            monthly_net = monthly_gross - (monthly_deductions + item_line_tax / 12) * float(day_count) / 30.0
                        else:
                            monthly_gross = item_line_earnings_nonstd + item_line_earnings_standard / 12
                            monthly_deductions = item_line_deductions_nonstd + item_line_deductions_standard / 12
                            monthly_net = monthly_gross - (monthly_deductions + item_line_tax / 12)
                        
                        ##_logger2.info("Gross=%f", item_line_gross)
                        ##_logger2.info("Net=%f", item_line_net)
                        ##_logger2.info("Relief=%f", item_line_relief)
                        ##_logger2.info("Taxable=%f", item_line_taxable)
                        ##_logger2.info("PAYE=%f", item_line_tax)
                        
                        #TODO Item list details are lost when persisted.
                        item_dict.update({'item_line_ids':item_list_lines})
                        if monthly_net > 0:
                            total_gross = total_gross + monthly_gross
                            total_net = total_net + monthly_net
                            if item_line_tax > 0:
                                total_tax = total_tax + (item_line_tax / 12)
            
                            item_dict.update({'gross_income':monthly_gross,'net_income':monthly_net,'balance_income':monthly_net,'taxable_income':(item_line_taxable / 12),'paye_tax':(item_line_tax / 12),'leave_allowance':item_line_leave,'active':emp.active,'retiring':item_line_retiring})
            
                            if not dept_summary.has_key(emp.department_id.id):
                                dept_summary[emp.department_id.id] = {'department_id':emp.department_id.id,'payroll_id':self.id,'total_taxable_income':0,'total_gross_income':0,'total_net_income':0,'total_paye_tax':0,'total_leave_allowance':0}

                            dept_summary[emp.department_id.id]['total_taxable_income'] += (item_line_taxable / 12)
                            dept_summary[emp.department_id.id]['total_gross_income'] += (item_line_gross / 12)
                            dept_summary[emp.department_id.id]['total_net_income'] += (item_line_net / 12)
                            dept_summary[emp.department_id.id]['total_paye_tax'] += (item_line_tax / 12)
                            dept_summary[emp.department_id.id]['total_leave_allowance'] += item_line_leave
                        else:
                            item_dict.update({'gross_income':0,'net_income':0,'balance_income':0,'taxable_income':0,'paye_tax':0,'leave_allowance':0,'active':emp.active,'retiring':item_line_retiring,'resolve':True})
                            
                        item_list.append(item_dict)
                    
                    self.payroll_item_ids = item_list
                    self.total_net_payroll = total_net
                    self.total_balance_payroll = total_net
                    self.total_gross_payroll = total_gross
                    self.total_taxable_payroll = total_taxable
                    self.total_tax_payroll = total_tax
                    self.processing_time = (time.time() - tic)
                    self.payroll_summary_ids = dept_summary.values()
            
                    #Process subventions
                    subvention_list = []
                    for subv in subventions:
                        subvention_list.append({'department_id': subv.org_id.id,'name': subv.name,'active': subv.active,'amount': subv.amount,'payroll_id':self.id})
                    self.subvention_item_ids = subvention_list
                    
                if self.do_pension:                    
                    tic = time.time()
                    total_gross = 0
                    total_net = 0
                    
                    #List all pension deductions
                    deductions_pension = self.env['ng.state.payroll.deduction.pension'].search([('active', '=', 'True')])
                    _logger.info("Count deductions_pension= %d", len(deductions_pension))        
        
                    pensioners = self.env['hr.employee'].search([('active', '=', 'True'), ('status_id.name', '=', 'PENSIONED')])
                    _logger.info("Count pensioners= %d", len(pensioners))        
                    
                    item_list = []
    
                    for pen in pensioners:
                        item_list_lines = []                
                        pension_amount = pen.annual_pension / 12
                        gross_amount = pension_amount
                        active_flag = False
                        if pen.status_id.name == 'PENSIONED':
                            active_flag = True
                        item_dict = {'employee_id':pen.id, 'active':active_flag, 'payroll_id':self.id, 'gross_income':pension_amount}
                        ded_amount = 0
                        item_list_lines.append({'name':'Monthly Pension', 'amount':pension_amount})
                        for ded in deductions_pension:
                            if len(ded.whitelist_ids) > 0 and pen in ded.whitelist_ids:
                                if not ded.fixed:
                                    ded_amount = pension_amount * ded.amount / 100
                                else:
                                    ded_amount = ded.amount
                            if not pen in ded.blacklist_ids:
                                if not ded.fixed:
                                    ded_amount = pension_amount * ded.amount / 100
                                else:
                                    ded_amount = ded.amount
                            pension_amount -= ded_amount
                            item_list_lines.append({'name':ded.name, 'amount':-ded_amount})
                        total_gross = total_gross + gross_amount
                        total_net = total_net + pension_amount                            
                        #TODO Item list details are lost when persisted.
                        item_dict.update({'item_line_ids':item_list_lines, 'gross_income':gross_amount, 'net_income':pension_amount, 'balance_income':pension_amount})
                        item_list.append(item_dict)
                    
                    self.pension_item_ids = item_list                                    
                    self.total_net_pension = total_net
                    self.total_balance_pension = total_net
                    self.total_gross_pension = total_gross
                    self.processing_time = (time.time() - tic)
                
        if (self.do_payroll or self.do_pension) and self.state == 'in_progress':
            self.update({'state':prev_state})
            
    @api.onchange('do_dry_run')
    def dry_run2(self):
        _logger.info("Calling dry_run2...state = %s", self.state)
        if self.in_progress:
            raise osv.except_osv(_('Info'), _('Processing already in progress.'))
                    
        if not self.state == 'in_progress':        
            if self.calendar_id and self.state == 'draft':
                if self.do_payroll:
                    tic = time.time()
                    item_list = []
                            
                    #List all tax rules
                    paye_taxrules = self.env['ng.state.payroll.taxrule'].search([('active', '=', 'True')])
                    
                    #Fetch all active employees *TODO* (and non-suspended employees)
                    employees = self.env['hr.employee'].search([('active', '=', 'True'), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id')
                    #employees = self.env['hr.employee'].search([('employee_no', 'in', ['65754','21978','37449','20925','11355','4599','22560','17771','4414','44552','30801','35323','32546','46842','31692','14867','58081','38892','10486','43932','40052','41927','32342','12507','7037','33280','41948','41944','17966','33498','32850','41969','79972','44957','16617','41520','36962','40917','47375','39139','41954','75063','32117','44436','41641','35637','44032','27535','78315','89000','73796','72297','88006','33686','31971','72773','72759','17781','17006','47208','65679','47929','87878','5107','5113','5084','05147_D','10868','12463','74083','84370','35370','6122','6123','6119','6148','1980','1936','1957','12460','33746','32017','75184','PP199','12026','12461','89040','86900','86904','86899','86897','86887','86889','86905','86902','86903','78632','34635','34700','34581','33514','34580','10616','4196','4153','4173','78404','73200','72110','72083','12090','40435','41637','3659','89014','5102','4146','4115','4144','4184','4075','4160','38250','44058','40933','10573','6449','5019','44071','43900','43905','43904','3654','3662','32454','33770','34609','31901','36897','34512','34479','31578','36384','34608','34489','30538','34835','38225','41797','88916','88893','86268','70972','44930','44055','PP196','PP197','PP198','78219','87195','87121','88751','88862','87056','88752','78298','87028','87196','87200','87217','88806','88396','78105','78167','88791','78265','86574','86591','88567','88662','88622','88568','78322','78324','78323','87036','87046','84001','88812','88813','78120','78119','78337','78209','88698','86461','86623','86621','86847','88378','78218','86525','86522','88366','86592','5099','5103','5078','74728','4182','41087','42024','72001','86028','36309','35622','86160','78217','87119','87183','87110','87142','88740','87067','88861','87097','88746','88738','87199','87118','87136','87108','87104','87211','87141','88674','88402','88403','88643','88520','88641','88486','88497','86520','86575','88481','86521','86506','88552','86515','88526','88387','88386','86236','87209','87206','78196','78184','87155','78197','78200','87022','78125','87124','88869','88796','87205','78124','87203','88736','87191','78310','78145','78144','88795','78134','88808','88850','78140','78135','78147','78148','78306','78253','78250','78252','78266','88734','88877','78171','87077','87081','78237','78228','88864','78172','78256','78247','87083','87087','87088','78244','87082','78233','88711','88684','86576','88710','86517','88547','88314','88658','88487','86518','88329','88652','86465','88599','86485','86486','88483','88610','88612','86484','88680','86494','86467','86512','86593','86012','78364','86251','86011','88917','88991','86242','86237','86267','88008','78389','88499','88642','88644','86587','86483','88544','88562','88659','88929','86273','86293','87030','78081','87149','87150','87147','78082','78121','78117','87134','78080','88857','87094','87044','87033','87055','87042','78112','78106','88866','78113','78111','88786','78341','87076','78343','88510','88664','88651','88450','88608','88470','88511','86510','86513','88513','86490','86504','88522','88976','86255','86247','88985','86294','86252','86222','86239','86269','88007','87176','86853','86852','78204','88872','88873','88797','87140','78215','78213','87175','87174','86250','86135','86285','86202','86171','86254','78386','86095','86162','86323','88712','88706','88654','88488','88576','88293','88705','88704','88655','88975','86168','86078','86169','86185','86024','86052','78414','86138','86136','86305','86300','87169','87168','86843','86842','86275','86627','86620','86141','86089','86292','86302','86638','86612','86622','86856','86860','86846','86316','86473','86629','86284','89051','89269','87171','87166','86240','86819','88673','89267','86848','88479','86260','86023','86201','88003','78246','86161','86244','78388','88896','88894','86155','78206','88731','78153','78377','88753','88019','88750','86079','86034','88913','86125','86087','88930','86199','88895','86499','78293','88732','86180','86597','86468','86523','88582','88860','86179','86139','86073','86235','86033','86154','86231','86232','88020','86322','86845','78395','86590','86464','88766','86571','86498','86812','86818','78270','86480','87072','87074','40043','67666','67670','72129','79905','67747','71814','72133','70975','66990','38218','38216','38219','42285','5100','5130','5068','5083','5098','5039','10598','10599','4150','40151','39322','89314','84071','84072','84564','84062','89860','89865','71444','71484','71486','88210','71482','71479','71446','71447','71445','71483','71485','71481','71480','71487','5049','75159','87557','78745','79954','70254','66996','66992','87564','67751','67669','71815','72141','71811','67654','72089','72021','68494_D','71998','75157','78757','44813','38959','39170','74692','74687','70580','70590','70591','74839','74701','68387','74696','78663','78664','75026','72086','45697','45688','65689','58514','41514','36320','30994','74675','44786','72816','41655','44639','58065','40542','38913','42538','27711','38455','25144','38560','41916','41363','7871','38660','121','6340','43641','42859','42936','6289','6299','6344'])])

                    _logger.info("Count employees= %d", len(employees))        
                    
                    total_gross = 0
                    total_net = 0
                    total_tax = 0
                    total_taxable = 0
                    
                    dept_summary = {}
            
                    for emp in employees:
                        #TODO When an employee has been reinstated in this calendar period,- 
                        #pick all previously inactive payroll items from previous calendar- 
                        #periods from the suspension month to current calendar period, move-
                        #them to current pay period and set them active
                        #_logger2.info("---------------------------------------------")
                        #_logger2.info("Name=%s", emp.name_related)
                        item_list_lines = []
                            
                        #Create Payroll Item and Payroll Item Lines
                        active_flag = False
                        if emp.status_id.name == 'ACTIVE':
                            active_flag = True
                        item_dict = {'employee_id':emp.id, 'active': active_flag, 'payroll_id': self.id}
                        item_line_gross = 0
                        item_line_leave = 0
                        item_line_deduction = 0
                        item_line_taxable = 0
                        item_line_relief = 200000
                        
                        #Calculate each standard earning
                        #for o in emp.standard_earnings.filtered(lambda r: r.active == True):
                        for o in emp.employee_earnings.filtered(lambda r: r.active == True):
                            amount = 0
                            if o.fixed:
                                amount = o.amount
                            else:
                                amount = o.amount * o.derived_from.amount * 0.01
                            item_line_gross += amount
                            #_logger2.info("Standard Earning[%s]=%f", o.name, amount)
                            item_list_lines.append({'name':o.name, 'amount':(amount / 12)})
                         
                        #Calculate each standard deduction
                        #for o in emp.standard_deductions.filtered(lambda r: r.active == True):
                        for o in emp.employee_deductions.filtered(lambda r: r.active == True):
                            amount = 0
                            if o.fixed:
                                amount = -o.amount
                            else:
                                if o.derived_from.fixed:
                                    amount = o.amount * o.derived_from.amount * 0.01
                                else:
                                    amount = o.amount * (o.derived_from.amount * 0.01 * o.derived_from.derived_from.amount) * 0.01
                            #_logger2.info("Standard Deduction[%s]=%f", o.name, amount)
                            item_line_deduction += -amount
                            item_list_lines.append({'name':o.name, 'amount':(amount / 12)})

                        #_logger2.info("Gross Income=%f", item_line_gross)
                        
                        percent_1percent = item_line_gross * 0.01
                        if percent_1percent > 200000:
                            item_line_relief = percent_1percent
                        item_line_relief += (item_line_gross * 0.2) #CRA relief
                        item_line_taxable = item_line_gross - item_line_deduction - item_line_relief
                        #Calculate PAYE Tax for each employee based on each taxable income items.
                        #TODO item_line_deduction should actually be reliefs - NHF, Pension, Party
                        total_taxable += (item_line_taxable / 12)
                        item_line_tax = 0
                        prev_to_amount = 0
                        for taxrule in paye_taxrules:
                            if item_line_taxable - taxrule.to_amount >= 0:
                                item_line_tax += ((taxrule.percentage / 100) * (taxrule.to_amount - prev_to_amount))
                                #_logger2.info("Amount=%f,Percentage=%f, PAYE=%f", (taxrule.to_amount - prev_to_amount), taxrule.percentage, item_line_tax)
                                prev_to_amount = taxrule.to_amount
                            else:
                                item_line_tax += ((taxrule.percentage / 100) * (item_line_taxable - prev_to_amount))
                                #_logger2.info("Amount=%f,Percentage=%f, PAYE=%f", (item_line_taxable - prev_to_amount), taxrule.percentage, item_line_tax)
                                break
                       
                        #Apply 1% PAYE rule
                        tax_1percent = item_line_taxable * 0.01
                        if item_line_tax < tax_1percent:
                            item_line_tax = tax_1percent                        
                        
                        item_line_net = item_line_gross - item_line_deduction - item_line_tax
                        
                        #_logger2.info("Gross=%f", item_line_gross)
                        #_logger2.info("Net=%f", item_line_net)
                        #_logger2.info("Taxable=%f", item_line_taxable)
                        #_logger2.info("PAYE=%f", item_line_tax)
                        
                        #TODO Item list details are lost when persisted.
                        item_dict.update({'item_line_ids':item_list_lines})
                        if item_line_gross > 0:
                            total_gross = total_gross + (item_line_gross / 12)
                            total_net = total_net + (item_line_net / 12)
                            total_tax = total_tax + (item_line_tax / 12)
            
                            item_dict.update({'gross_income':(item_line_gross / 12),'net_income':(item_line_net / 12),'balance_income':(item_line_net / 12),'taxable_income':(item_line_taxable / 12),'paye_tax':(item_line_tax / 12),'leave_allowance':item_line_leave,'active':emp.active})
            
                            if not dept_summary.has_key(emp.department_id.id):
                                dept_summary[emp.department_id.id] = {'department_id':emp.department_id.id,'payroll_id':self.id,'total_taxable_income':0,'total_gross_income':0,'total_net_income':0,'total_paye_tax':0,'total_leave_allowance':0}

                            dept_summary[emp.department_id.id]['total_taxable_income'] += (item_line_taxable / 12)
                            dept_summary[emp.department_id.id]['total_gross_income'] += (item_line_gross / 12)
                            dept_summary[emp.department_id.id]['total_net_income'] += (item_line_net / 12)
                            dept_summary[emp.department_id.id]['total_paye_tax'] += (item_line_tax / 12)
                            dept_summary[emp.department_id.id]['total_leave_allowance'] += item_line_leave
                        else:
                            item_dict.update({'gross_income':(item_line_gross / 12),'net_income':(item_line_net / 12),'balance_income':(item_line_net / 12),'taxable_income':(item_line_taxable / 12),'paye_tax':(item_line_tax / 12),'leave_allowance':item_line_leave,'active':emp.active,'resolve':True})
                            
                        item_list.append(item_dict)
                    
                    self.payroll_item_ids = item_list
                    self.total_net_payroll = total_net
                    self.total_balance_payroll = total_net
                    self.total_gross_payroll = total_gross
                    self.total_taxable_payroll = total_taxable
                    self.total_tax_payroll = total_tax
                    self.processing_time_payroll = (time.time() - tic)
                    self.payroll_summary_ids = dept_summary.values()
                    self.state = 'processed'
                    
                if self.do_pension:                    
                    tic = time.time()
                    total_gross = 0
                    total_net = 0
                    
                    #List all pension deductions
                    deductions_pension = self.env['ng.state.payroll.deduction.pension'].search([('active', '=', 'True')])
                    _logger.info("Count deductions_pension= %d", len(deductions_pension))        
        
                    pensioners = self.env['hr.employee'].search([('active', '=', 'True'), ('status_id.name', '=', 'PENSIONED')])
                    _logger.info("Count pensioners= %d", len(pensioners))        
                    
                    item_list = []
    
                    for pen in pensioners:
                        item_list_lines = []                
                        pension_amount = pen.annual_pension / 12
                        gross_amount = pension_amount
                        active_flag = False
                        if pen.status_id.name == 'PENSIONED':
                            active_flag = True
                        item_dict = {'employee_id':pen.id, 'active':active_flag, 'payroll_id':self.id, 'gross_income':pension_amount}
                        ded_amount = 0
                        item_list_lines.append({'name':'Monthly Pension', 'amount':pension_amount})
                        for ded in deductions_pension:
                            if len(ded.whitelist_ids) > 0 and pen in ded.whitelist_ids:
                                if not ded.fixed:
                                    ded_amount = pension_amount * ded.amount / 100
                                else:
                                    ded_amount = ded.amount
                            if not pen in ded.blacklist_ids:
                                if not ded.fixed:
                                    ded_amount = pension_amount * ded.amount / 100
                                else:
                                    ded_amount = ded.amount
                            pension_amount -= ded_amount
                            item_list_lines.append({'name':ded.name, 'amount':-ded_amount})
                        total_gross = total_gross + gross_amount
                        total_net = total_net + pension_amount                            
                        #TODO Item list details are lost when persisted.
                        item_dict.update({'item_line_ids':item_list_lines, 'gross_income':gross_amount, 'net_income':pension_amount, 'balance_income':pension_amount})
                        item_list.append(item_dict)
                    
                    self.pension_item_ids = item_list                                    
                    self.total_net_pension = total_net
                    self.total_balance_pension = total_net
                    self.total_gross_pension = total_gross
                    self.processing_time_pension = (time.time() - tic)
                    self.state = 'processed'
     
    def try_generate_reports(self, cr, uid, context=None):
        _logger.info("Running try_generate_reports cron-job...")
        payroll_singleton = self.pool.get('ng.state.payroll.payroll')
        payroll_ids = payroll_singleton.search(cr, uid, [('state', '=', 'closed'), ('generate_reports', '=', True)], context=context)
        for p in payroll_singleton.browse(cr, uid, payroll_ids, context=context):
            p.process_reports()

        return True
       
    @api.multi
    def process_reports(self):
        _logger.info("process_reports : processing payroll %s -> %s", self.name, self.mda_emails)
        if self.mda_emails:
            path = '/tmp/payroll_reports_' + str(self.id)
            if not os.path.exists(path):
                os.makedirs(path)
    
            _logger.info("process_reports : summary_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_summary_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/summary_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            if self.do_pension:
                _logger.info("process_reports : pension_exec_summary_report")
                file_data = BytesIO()
                workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
                xlsx_data = pension_exec_summary_rep.generate_xlsx_report(workbook, {}, self, file_data)
                fd = open(path + '/pension_exec_summary_report.xlsx', 'w')
                fd.write(xlsx_data)
            
            _logger.info("process_reports : payroll_exec_summary_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_exec_summary_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/payroll_exec_summary_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : exec_summary2_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_exec_summary2_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/exec_summary2_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : paye_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_paye_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/paye_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : item_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_item_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/item_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : tescom_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_tescom_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/tescom_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : tescom_school_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_tescom_school_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/tescom_school_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : leavebonus_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_leavebonus_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/leavebonus_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : mda_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_mda_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/mda_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : tco_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = pension_tco_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/tco_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : mda_deduction_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_mda_deduction_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/mda_deduction_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : mda_deduction_head_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_mda_deduction_head_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/mda_deduction_head_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : mda_summarized_report")
            file_data = BytesIO()
            workbook = xlsxwriter.Workbook(file_data, {'constant_memory': False})
            xlsx_data = payroll_mda_summarized_rep.generate_xlsx_report(workbook, {}, self, file_data)
            fd = open(path + '/mda_summarized_report.xlsx', 'w')
            fd.write(xlsx_data)
            
            _logger.info("process_reports : zipping...")
            shutil.make_archive(path, 'zip', path)

            _logger.info("process_reports : mailing...")
            receivers = self.mda_emails #Comma separated email addresses
            message = "Dear Sir,\nPlease find the reports for the payroll as found in the title of this email.\n\nThank you.\n"
            smtp_obj = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
            smtp_obj.ehlo()
            #smtp_obj.starttls()
            #smtp_obj.ehlo()
            smtp_obj.login(user="osun.payroll@gmail.com", password="p@55w0rd1939")
            sender = 'osun.payroll@gmail.com'
            msg = MIMEMultipart()
            msg['Subject'] = 'Payroll Closed - ' + self.name 
            msg['From'] = sender
            #msg['To'] = ', '.join(receivers)
            msg['To'] = receivers
                             
            part = MIMEBase('application', "octet-stream")
            part.set_payload(open(path + '.zip', "rb").read())
            Encoders.encode_base64(part)                            
            part.add_header('Content-Disposition', 'attachment; filename="payroll_reports_' + str(self.id) + '.zip"')
            msg.attach(MIMEText(message))
            msg.attach(part)
            smtp_obj.sendmail(sender, receivers, msg.as_string())
            self.write({'generate_reports':False})
            _logger.info("process_reports : report mailed.")           
        
class ng_state_payroll_promotion(models.Model):
    '''
    Employee Promotion
    '''
    _name = "ng.state.payroll.promotion"
    _description = 'Employee Promotion'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'batch_number': fields.char('Batch Group', help='Batch Group; for batch approval', required=False),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'promotion_rule_id': fields.many2one('ng.state.payroll.promotion.rule', 'Promotion Rule'),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('pending', 'Pending'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'promotion_type': fields.selection([
            ('auto', 'Automatic'),
            ('manual', 'Manual'),
        ], 'Type', readonly=True),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'end_date': fields.date('End Date', required=False, readonly=True, states={'draft': [('readonly', False)]}),
        'acting': fields.boolean('Acting', help='Tick check-box if the employee is promoted in acting capacity'),
        'error_msg': fields.char('Error Message', help='Error Message holding up process', required=False),
        'from_pay_category_id': fields.many2one('ng.state.payroll.paycategory', 'From Pay Category'),
        'from_pay_scheme_id': fields.many2one('ng.state.payroll.payscheme', 'From Pay Scheme'),
        'from_level_id': fields.many2one('ng.state.payroll.level', 'From Grade'),
        'from_grade_level': fields.selection([
            (1, 'GL-1'),
            (2, 'GL-2'),
            (3, 'GL-3'),
            (4, 'GL-4'),
            (5, 'GL-5'),
            (6, 'GL-6'),
            (7, 'GL-7'),
            (8, 'GL-8'),
            (9, 'GL-9'),
            (10, 'GL-10'),
            (11, 'GL-11'),
            (12, 'GL-12'),
            (13, 'GL-13'),
            (14, 'GL-14'),
            (15, 'GL-15'),
            (16, 'GL-16'),
            (17, 'GL-17'),
            (18, 'GL-18'),
            (19, 'GL-19'),
            (20, 'GL-20'),
        ], 'From Grade Level'),
        'from_grade_step': fields.selection([
            (1, 'Step-1'),
            (2, 'Step-2'),
            (3, 'Step-3'),
            (4, 'Step-4'),
            (5, 'Step-5'),
            (6, 'Step-6'),
            (7, 'Step-7'),
            (8, 'Step-8'),
            (9, 'Step-9'),
            (10, 'Step-10'),
            (11, 'Step-11'),
            (12, 'Step-12'),
            (13, 'Step-13'),
            (14, 'Step-14'),
            (15, 'Step-15'),
            (16, 'Step-16'),
            (17, 'Step-17'),
            (18, 'Step-18'),
            (19, 'Step-19'),
            (20, 'Step-20'),
        ], 'From Grade Step'),                
        'to_pay_category_id': fields.many2one('ng.state.payroll.paycategory', 'To Pay Category', required=True),
        'to_pay_scheme_id': fields.many2one('ng.state.payroll.payscheme', 'To Pay Scheme', required=True),
        'to_level_id': fields.many2one('ng.state.payroll.level', 'To Grade'),
        'to_grade_level': fields.selection([
            (1, 'GL-1'),
            (2, 'GL-2'),
            (3, 'GL-3'),
            (4, 'GL-4'),
            (5, 'GL-5'),
            (6, 'GL-6'),
            (7, 'GL-7'),
            (8, 'GL-8'),
            (9, 'GL-9'),
            (10, 'GL-10'),
            (12, 'GL-12'),
            (13, 'GL-13'),
            (14, 'GL-14'),
            (15, 'GL-15'),
            (16, 'GL-16'),
            (17, 'GL-17'),
            (18, 'GL-18'),
            (19, 'GL-19'),
            (20, 'GL-20'),
        ], 'To GL'),
        'to_grade_step': fields.selection([
            (1, 'Step-1'),
            (2, 'Step-2'),
            (3, 'Step-3'),
            (4, 'Step-4'),
            (5, 'Step-5'),
            (6, 'Step-6'),
            (7, 'Step-7'),
            (8, 'Step-8'),
            (9, 'Step-9'),
            (10, 'Step-10'),
            (11, 'Step-11'),
            (12, 'Step-12'),
            (13, 'Step-13'),
            (14, 'Step-14'),
            (15, 'Step-15'),
            (16, 'Step-16'),
            (17, 'Step-17'),
            (18, 'Step-18'),
            (19, 'Step-19'),
            (20, 'Step-20'),
        ], 'To Step'),
    }

    _rec_name = 'date'
        
    @api.onchange('from_pay_scheme_id')
    def from_level_id_update(self):
        return {'domain': {'from_level_id': [('paygrade_id.payscheme_id','=',self.from_pay_scheme_id.id)] }}
        
    @api.onchange('to_pay_scheme_id')
    def to_level_id_update(self):
        return {'domain': {'to_level_id': [('paygrade_id.payscheme_id','=',self.to_pay_scheme_id.id)] }}
    
    @api.onchange('date')
    def check_effective_date(self):
        if self.date >= self.employee_id.retirement_due_date:    
            raise osv.except_osv(
                _('Warning!'),
                _('The promotion date is on or before the employee retirement date.')
            )
            
    @api.multi
    def confirm(self):
        _logger.info("confirm - %s", 'confirm')
        self.write({'state':'confirm'})
     
    @api.multi
    def cancel(self):
        _logger.info("cancel - %s", 'cancel')
        self.write({'state':'cancel'})
            
    @api.multi
    def pending(self):
        _logger.info("pending - %s", 'pending')
        self.write({'state':'pending'})
   
    @api.model
    def create(self, vals):
        employee_obj = self.env['hr.employee']
        this_employee = employee_obj.browse(vals.get('employee_id'))
        vals.update({
            'from_pay_category_id':this_employee.paycategory_id.id,
            'from_pay_scheme_id':this_employee.payscheme_id.id,
            'from_level_id':this_employee.level_id.id,
            'from_grade_step':this_employee.level_id.step,
            'from_grade_level':this_employee.level_id.paygrade_id.level
        })
        res = super(ng_state_payroll_promotion, self).create(vals)
            
        return res   
   
    def _get_default_domain_employees(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        employee_obj = self.pool.get('hr.employee')

        this_user = users_obj.browse(cr, uid, uid, context=context)
        employees = []
        if this_user.domain_mdas:
            employees = employee_obj.search(cr, uid, [('department_id.id', 'in', this_user.domain_mdas.ids)], context=context)
        else:
            employees = employee_obj.search(cr, uid, [], context=context)

        return employees
    
    _defaults = {
        'employee_id': _get_default_domain_employees,
        'state': 'draft',
        'promotion_type': 'manual'
    }

    _track = {
        'state': {
            'ng_state_payroll_promotion.mt_alert_promo_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_promotion.mt_alert_promo_pending':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'pending',
            'ng_state_payroll_promotion.mt_alert_promo_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }

    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            domain = [('state', '=', 'confirm')]
            return domain

        return False

    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Promotion!'),
                    _('Promotion has been initiated. Either cancel the promotion or create another promotion to undo it.')
                )

        return super(ng_state_payroll_promotion, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):
        today = datetime.now().date()
        for item_obj in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                item_obj.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def _check_state(self, cr, uid, employee_id, effective_date, context=None):
        _logger.info("_check_state - %d", employee_id)
        employee_obj = self.pool.get('hr.employee')
        data = employee_obj.read(
            cr, uid, employee_id, ['state', 'retirement_due_date'], context=context) 
        if data.get('retirement_due_date', False) and data['retirement_due_date'] != '':
            retirementDate = datetime.strptime(
                data['retirement_due_date'], DEFAULT_SERVER_DATE_FORMAT)
            dEffective = datetime.strptime(
                effective_date, DEFAULT_SERVER_DATE_FORMAT)
            if dEffective >= retirementDate:
                self.write(cr, uid, self.id, {'error_msg': 'Effective Date cannot be after Retirement Due Date.'}, context=context)
                return False
                
        return True

    def promotion_state_confirm(self, cr, uid, ids, context=None):
        for promo in self.browse(cr, uid, ids, context=context):
            _logger.info("before state_confirm - %d", uid)
            if self._check_state(
                cr, uid, promo.employee_id.id, promo.date, context=context):
                self.write(cr, uid, promo.id, {'state': 'confirm'}, context=context)
            _logger.info("after state_confirm - %d", uid)

        return True

    def promotion_state_done(self, cr, uid, ids, context=None):
        employee_obj = self.pool.get('hr.employee')
        gradelevel_obj = self.pool.get('ng.state.payroll.level')
        cron_obj = self.pool.get('ir.cron')
        
        today = datetime.now().date()

        for promo in self.browse(cr, uid, ids, context=context):
            if datetime.strptime(
                promo.date, DEFAULT_SERVER_DATE_FORMAT
            ).date() <= today and promo.state == 'pending':
                if self._check_state(
                    cr, uid, promo.employee_id.id, promo.date,
                    context=context):
                    #level_ids = gradelevel_obj.search(cr, uid, [('active', '=', True), ('step', '=', promo.to_grade_step), ('paygrade_id.level', '=', promo.to_grade_level)], context=context)
                    #if level_ids:
                    employee_obj.write(
                        cr, uid, promo.employee_id.id, {
                            'resolved_earn_dedt': False,
                            'last_promotion_date': promo.date,
                            'paycategory_id': promo.to_pay_category_id.id,
                            'payscheme_id': promo.to_pay_scheme_id.id,
                            'grade_step':promo.to_level_id.step,
                            'grade_level': promo.to_level_id.paygrade_id.level,
                            'level_id':promo.to_level_id.id},
                        context=context)
                    hrevent_obj = self.pool.get('ng.state.payroll.hrevent')
                    hrevent_obj.create(cr, uid, {'employee_id':promo.employee_id.id, 'activity_type':'promotion', 'activity_id':promo.id})
                    cron_ids = cron_obj.search(cr, uid, [('name', '=', 'Resolve Standard Earnings and Deductions')], context=context)
                    cron_rec = cron_obj.browse(cr, uid, cron_ids[0], context=context)
                    nextcall = datetime.now() + timedelta(seconds=5)
                    cron_rec.write({'nextcall':nextcall.strftime(DEFAULT_SERVER_DATETIME_FORMAT)})
                    self.write(
                        cr, uid, promo.id, {'state': 'approved'}, context=context)
            else:
                return False

        return True

    def try_pending_promotions(self, cr, uid, context=None):
        """Completes pending promotions. Called from
        the scheduler."""
        
        _logger.info("Running try_pending_promotions cron-job...")
        
        promo_obj = self.pool.get('ng.state.payroll.promotion')
        today = datetime.now().date()
        promo_ids = promo_obj.search(cr, uid, [
            ('state', '=', 'pending'),
            ('date', '<=', today.strftime(
                DEFAULT_SERVER_DATE_FORMAT)),
        ], context=context)
        
        self.promotion_state_done(cr, uid, promo_ids, context=context)

        return True

    def try_init_next_promotion_dates(self, cr, uid, context=None):
        """Initializes next promotion dates when blank."""

        _logger.info("Running try_init_next_promotion_dates cron-job...")
        
        #Fetch all active employees due for a promotion
        employee_obj = self.pool.get('hr.employee')
        promo_rule_obj = self.pool.get('ng.state.payroll.promotion.rule')
        employee_ids = employee_obj.search(cr, uid, [('active', '=', True), ('status_id.name', '=', 'ACTIVE'), ('next_promotion_date', '=', False)], context=context)        

        _logger.info("try_init_next_promotion_dates - employees=%d", len(employee_ids))
        for emp in employee_obj.browse(cr, uid, employee_ids, context=context):
            promo_rule_ids = promo_rule_obj.search(cr, uid, [('from_grade_level', '=', emp.grade_level)], context=context)
            if len(promo_rule_ids) == 1:
                promo_rules = promo_rule_obj.browse(cr, uid, promo_rule_ids[0], context=context)
                if emp.last_promotion_date:        
                    next_promo_date = datetime.strptime(emp.last_promotion_date, DEFAULT_SERVER_DATE_FORMAT) + relativedelta(years=promo_rules.years_due)
                    today = date.today()
                    if next_promo_date <= datetime(today.year, today.month, today.day):
                        promo_rules2_ids = promo_rule_obj.search(cr, uid, [('from_grade_level', '=', (emp.grade_level + 1))], context=context)
                        if len(promo_rules2_ids) == 1:
                            promo_rules2 = promo_rule_obj.browse(cr, uid, promo_rules2_ids[0], context=context)  
                            end_date = date.today() + relativedelta(years=promo_rules2.years_due)
                            cr.execute("update hr_employee set next_promotion_date='" + end_date.strftime(DEFAULT_SERVER_DATE_FORMAT) + "' where id=" + str(emp.id))
        cr.commit()
        return True

    def try_due_promotions(self, cr, uid, context=None):
        """Creates automatic promotions for confirmation."""

        _logger.info("Running try_due_promotions cron-job...")
        
        #Fetch all active employees due for a promotion
        today = datetime.now().date()
        employee_obj = self.pool.get('hr.employee')
        promo_rule_obj = self.pool.get('ng.state.payroll.promotion.rule')
        promo_obj = self.pool.get('ng.state.payroll.promotion')
        employee_ids = employee_obj.search(cr, uid, [('active', '=', True), ('next_promotion_date', '!=', False), ('status_id.name', '=', 'ACTIVE'), ('next_promotion_date', '<=', today.strftime(DEFAULT_SERVER_DATE_FORMAT))], order='id')        

        _logger.info("try_due_promotions - employees=%d", len(employee_ids))
        for emp in employee_obj.browse(cr, uid, employee_ids, context=context):
            promo_rule_ids = promo_rule_obj.search(cr, uid, [('from_grade_level', '=', (emp.grade_level + 1))])
            if len(promo_rule_ids) == 1:
                promo_rules = promo_rule_obj.browse(cr, uid, promo_rule_ids[0], context=context) 
                end_date = date.today() + relativedelta(years=promo_rules[0].years_due)
                next_level = emp.grade_level + 1
                if emp.grade_level == 10:
                    next_level = 12
                promo_obj.create(cr, uid, {
                    'employee_id':emp.id,
                    'promotion_rule_id':promo_rules[0].id,
                    'state':'confirm','promotion_type':'auto',
                    'date':today.strftime(DEFAULT_SERVER_DATE_FORMAT),
                    'end_date':end_date.strftime(DEFAULT_SERVER_DATE_FORMAT),
                    'from_pay_category_id':emp.paycategory_id.id,
                    'from_pay_scheme_id':emp.payscheme_id.id,
                    'from_grade_level':emp.grade_level,
                    'to_pay_category_id':emp.paycategory_id.id,
                    'to_pay_scheme_id':emp.payscheme_id.id,
                    'to_grade_level': next_level,
                }, context=context)
                employee_obj.write(cr, uid, emp.id, {'next_promotion_date':end_date.strftime(DEFAULT_SERVER_DATE_FORMAT)}, context=context)
        cr.commit()
        return True

    def onchange_employee(self, cr, uid, ids, employee_id, context=None):
        res = {'value': {'from_pay_category_id': False, 'from_pay_scheme_id': False, 'from_level_id': False, 'from_grade_level': False}}
        
        if employee_id:
            ee = self.pool.get('hr.employee').browse(
                cr, uid, employee_id, context=context)
            res['value']['from_pay_category_id'] = ee.paycategory_id.id
            res['value']['from_pay_scheme_id'] = ee.payscheme_id.id
            res['value']['from_level_id'] = ee.level_id.id
            res['value']['from_grade_step'] = ee.level_id.step
            res['value']['from_grade_level'] = ee.level_id.paygrade_id.level

        return res
                    
class ng_state_payroll_promotion_rule(models.Model):
    '''
    Promotion Rule
    '''
    _name = "ng.state.payroll.promotion.rule"
    _description = 'Promotion Rule'
    _columns = {    
        'from_grade_level': fields.selection([
            (1, 'GL-1'),
            (2, 'GL-2'),
            (3, 'GL-3'),
            (4, 'GL-4'),
            (5, 'GL-5'),
            (6, 'GL-6'),
            (7, 'GL-7'),
            (8, 'GL-8'),
            (9, 'GL-9'),
            (10, 'GL-10'),
            (12, 'GL-12'),
            (13, 'GL-13'),
            (14, 'GL-14'),
            (15, 'GL-15'),
            (16, 'GL-16'),
            (17, 'GL-17'),
            (18, 'GL-18'),
            (19, 'GL-19'),
            (20, 'GL-20'),
        ], 'From Grade Level'),
        'to_grade_level': fields.selection([
            (1, 'GL-1'),
            (2, 'GL-2'),
            (3, 'GL-3'),
            (4, 'GL-4'),
            (5, 'GL-5'),
            (6, 'GL-6'),
            (7, 'GL-7'),
            (8, 'GL-8'),
            (9, 'GL-9'),
            (10, 'GL-10'),
            (12, 'GL-12'),
            (13, 'GL-13'),
            (14, 'GL-14'),
            (15, 'GL-15'),
            (16, 'GL-16'),
            (17, 'GL-17'),
            (18, 'GL-18'),
            (19, 'GL-19'),
            (20, 'GL-20'),
        ], 'To Grade Level'),
        'years_due': fields.integer('Due on Years', help='Due on Years'),
    }
    
class ng_state_payroll_disciplinary(models.Model):
    '''
    Payroll Disciplinary (Suspension/Reinstatement)
    '''
    _name = "ng.state.payroll.disciplinary"
    _description = 'Payroll Disciplinary (Suspension/Reinstatement)'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'batch_number': fields.char('Batch Group', help='Batch Group; for batch approval', required=False),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('pending', 'Pending'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'action_type': fields.selection([
            ('suspension', 'Suspension'),
            ('reinstatement', 'Reinstatement'),
        ], 'Type', required=True, readonly=False),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'end_date': fields.date('End Date', required=False, readonly=True, states={'draft': [('readonly', False)]}),
        'unpaid_suspension': fields.boolean('Unpaid Suspension', help='If checked, employee is not paid during the suspension period.'),
        'pensioned': fields.boolean('Pensioner', help='Indicates that this applies to a pensioner if true.'),
    }

    _rec_name = 'date'
    
    @api.multi
    def confirm(self):
        _logger.info("confirm - %s", 'confirm')
        self.write({'state':'confirm'})
     
    @api.multi
    def cancel(self):
        _logger.info("cancel - %s", 'cancel')
        self.write({'state':'cancel'})
    
    @api.multi
    def pending(self):
        _logger.info("pending - %s", 'pending')
        self.write({'state':'pending'})
   
    def _get_default_domain_employees(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        employee_obj = self.pool.get('hr.employee')

        this_user = users_obj.browse(cr, uid, uid, context=context)
        employees = []
        if this_user.domain_mdas:
            employees = employee_obj.search(cr, uid, [('department_id.id', 'in', this_user.domain_mdas.ids)], context=context)
        else:
            employees = employee_obj.search(cr, uid, [], context=context)

        return employees
    
    _defaults = {
        'employee_id': _get_default_domain_employees,
        'state': 'draft',
        'action_type': 'suspension',
        'unpaid_suspension': False,
        'pensioned': False,
    }
       
    _track = {
        'state': {
            'ng_state_payroll_disciplinary.mt_alert_disc_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_disciplinary.mt_alert_disc_pending':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'pending',
            'ng_state_payroll_disciplinary.mt_alert_disc_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }

    def _check_state(self, cr, uid, employee_id, effective_date, context=None):
        _logger.info("_check_state - %d", employee_id)
        employee_obj = self.pool.get('hr.employee')
        data = employee_obj.read(
            cr, uid, employee_id, ['state', 'retirement_due_date'], context=context) 
        if data.get('retirement_due_date', False) and data['retirement_due_date'] != '':
            retirementDate = datetime.strptime(
                data['retirement_due_date'], DEFAULT_SERVER_DATE_FORMAT)
            dEffective = datetime.strptime(
                effective_date, DEFAULT_SERVER_DATE_FORMAT)
            if dEffective >= retirementDate:
                return False
                
        return True
    
    def _needaction_domain_get(self, cr, uid, context=None):

        users_obj = self.pool.get('res.users')

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            domain = [('state', '=', 'confirm')]
            return domain

        return False

    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Disciplinary action!'),
                    _('Disciplinary action has been initiated. Either cancel the disciplinary action or create another to undo it.')
                )

        return super(ng_state_payroll_disciplinary, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for disc in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                disc.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def disciplinary_state_confirm(self, cr, uid, ids, context=None):
        _logger.info("before state_confirm - %d", uid)
        for disc in self.browse(cr, uid, ids, context=context):
            if self._check_state(
                cr, uid, disc.employee_id.id, disc.date, context=context):
                self.write(cr, uid, disc.id, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def disciplinary_state_done(self, cr, uid, ids, context=None):

        employee_obj = self.pool.get('hr.employee')
        today = datetime.now().date()

        for disc in self.browse(cr, uid, ids, context=context):
            if datetime.strptime(
                disc.date, DEFAULT_SERVER_DATE_FORMAT
            ).date() <= today and disc.state == 'pending':
                if self._check_state(
                    cr, uid, disc.employee_id.id, disc.date,
                    context=context):
                    status_obj = self.pool.get('ng.state.payroll.status')
                    if disc.action_type == 'suspension':
                        suspended_status = False
                        if disc.pensioned:
                            suspended_status_ids = status_obj.search(cr, uid, [('name', '=', 'PENSIONED_SUSPENDED')], context=context)
                            suspended_status = status_obj.browse(cr, uid, suspended_status_ids[0], context=context)
                        else:
                            suspended_status_ids = status_obj.search(cr, uid, [('name', '=', 'SUSPENDED')], context=context)
                            suspended_status = status_obj.browse(cr, uid, suspended_status_ids[0], context=context)
                        employee_obj.write(
                            cr, uid, disc.employee_id.id, {
                                'status_id': suspended_status[0].id},
                            context=context)
                    else:
                        active_status = False
                        if disc.pensioned:
                            active_status_ids = status_obj.search(cr, uid, [('name', '=', 'PENSIONED')], context=context)
                            active_status = status_obj.browse(cr, uid, active_status_ids[0], context=context)
                        else:
                            active_status_ids = status_obj.search(cr, uid, [('name', '=', 'ACTIVE')], context=context)
                            active_status = status_obj.browse(cr, uid, active_status_ids[0], context=context)
                        employee_obj.write(
                            cr, uid, disc.employee_id.id, {
                                'status_id': active_status[0].id},
                            context=context)
                    self.write(
                        cr, uid, disc.id, {'state': 'approved'}, context=context)
                    hrevent_obj = self.pool.get('ng.state.payroll.hrevent')
                    hrevent_obj.create(cr, uid, {'employee_id':disc.employee_id.id, 'activity_type':disc.action_type, 'activity_id':disc.id})
            else:
                return False

        cr.commit()
        return True

    def try_pending_disciplinary_actions(self, cr, uid, context=None):
        """Completes pending disciplinary actions. Called from
        the scheduler."""
        
        _logger.info("Running try_pending_disciplinary_actions cron-job...")

        disc_obj = self.pool.get('ng.state.payroll.disciplinary')
        today = datetime.now().date()
        disc_ids = disc_obj.search(cr, uid, [
            ('state', '=', 'pending'),
            ('date', '<=', today.strftime(
                DEFAULT_SERVER_DATE_FORMAT)),
        ], context=context)
        
        self.disciplinary_state_done(cr, uid, disc_ids, context=context)

        return True
    
class ng_state_payroll_loan_payment(models.Model):
    '''
    Payroll Employee Payment
    '''
    _name = "ng.state.payroll.loan.payment"
    _description = 'Payroll Employee Loan Payment'
    
    _columns = {
        'loan_id': fields.many2one('ng.state.payroll.loan', 'Loan', required=True, readonly=True),
        'date': fields.date('Payment Date', required=True, readonly=True),
        'amount': fields.float('Paid Amount', required=True, readonly=True),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True, readonly=True),
        'payroll_id': fields.many2one('ng.state.payroll.payroll', 'Payroll', required=True, readonly=True),
    }
       
class ng_state_payroll_loan(models.Model):
    '''
    Payroll Employee Loan
    '''
    _name = "ng.state.payroll.loan"
    _description = 'Payroll Employee Loan'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'batch_number': fields.char('Batch Group', help='Batch Group; for batch approval', required=False),
        'name': fields.char('Name', help='Loan Name', required=True),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('pending', 'Pending'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'active': fields.boolean('Active', help='Active', required=True),
        'loan_amount': fields.float('Loan Amount', help='Loan Amount', required=True),
        'total_payment_amount': fields.float('Total Payment Amount', help='Total Payment Amount', readonly=True, required=True, compute='_payment_amount_update', states={'draft': [('readonly', False)]}),
        'payment_amount': fields.float('Payment Amount', help='Payment Amount', readonly=True, required=True, compute='_payment_amount_update', states={'draft': [('readonly', False)]}),
        'tenure': fields.integer('Tenure (Months)', help='Tenure (Months)', required=True),
        'interest_rate': fields.float('Interest Rate', help='Interest Rate', required=True),
        'payment_ids': fields.one2many('ng.state.payroll.loan.payment','loan_id','Loan Payments'),
        'employee_ids': fields.many2many('hr.employee', 'rel_employee_loan', 'loan_id', 'employee_id', 'Employees'),
    }
    
    _defaults = {
        'active': True,
        'tenure': 24,
        'interest_rate': 0.0,
    }     
 
    _rec_name = 'date'
        
    @api.depends('loan_amount', 'tenure', 'interest_rate')
    def _payment_amount_update(self):
        for loan in self:
            if loan.tenure <= 0:
                loan.tenure = 24
            if loan.interest_rate <= 0:
                loan.interest_rate = 0.0
            if loan.loan_amount < 0:
                loan.loan_amount = 0.0
            loan.total_payment_amount = loan.loan_amount * loan.interest_rate / 100 + loan.loan_amount
            loan.payment_amount = loan.total_payment_amount / loan.tenure
                
             
    @api.multi
    def confirm(self):
        _logger.info("confirm - %s", 'confirm')
        self.write({'state':'confirm'})
     
    @api.multi
    def cancel(self):
        _logger.info("cancel - %s", 'cancel')
        self.write({'state':'cancel'})
    
    @api.multi
    def pending(self):
        _logger.info("pending - %s", 'pending')
        self.write({'state':'pending'})
   
    def _get_default_domain_employees(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        employee_obj = self.pool.get('hr.employee')

        this_user = users_obj.browse(cr, uid, uid, context=context)
        employees = []
        if this_user.domain_mdas:
            employees = employee_obj.search(cr, uid, [('department_id.id', 'in', this_user.domain_mdas.ids)], context=context)
        else:
            employees = employee_obj.search(cr, uid, [], context=context)

        return employees
    
    _defaults = {
        'employee_id': _get_default_domain_employees,
        'state': 'draft',
    }
       
    _track = {
        'state': {
            'ng_state_payroll_loan.mt_alert_loan_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_loan.mt_alert_loan_pending':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'pending',
            'ng_state_payroll_loan.mt_alert_loan_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }

    def _check_state(self, cr, uid, effective_date, context=None):
        _logger.info("_check_state - %s", effective_date)
                
        return True
    
    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        _logger.info("_needaction_domain_get - %s", users_obj)

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            _logger.info("_needaction_domain_get - is HR Manager")
            domain = [('state', '=', 'confirm')]
            return domain

        return False

    
    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Loan action!'),
                    _('Loan action has been initiated. Either cancel the loan action or create another to undo it.')
                )

        return super(ng_state_payroll_loan, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for disc in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                disc.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def loan_state_confirm(self, cr, uid, ids, context=None):
        _logger.info("before state_confirm - %d", uid)
        for disc in self.browse(cr, uid, ids, context=context):
            if self._check_state(
                cr, uid, disc.date, context=context):
                self.write(cr, uid, disc.id, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def loan_state_done(self, cr, uid, ids, context=None):
        _logger.info("Calling loan_state_done...")
        today = datetime.now().date()

        for o in self.browse(cr, uid, ids, context=context):
            if datetime.strptime(
                o.date, DEFAULT_SERVER_DATE_FORMAT
            ).date() <= today and o.state == 'pending':
                if self._check_state(
                    cr, uid, o.date,
                    context=context):
                    self.write(cr, uid, o.id, {'state': 'approved'}, context=context)
                    hrevent_obj = self.pool.get('ng.state.payroll.hrevent')
                    hrevent_obj.create(cr, uid, {'employee_id':o.employee_id.id, 'activity_type':'loan', 'activity_id':o.id})
            else:
                return False
        cr.commit()
        return True

    def try_pending_loan_actions(self, cr, uid, context=None):
        """Completes pending loan actions. Called from
        the scheduler."""

        _logger.info("Running try_pending_loan_actions cron-job...")
        
        disc_obj = self.pool.get('ng.state.payroll.loan')
        today = datetime.now().date()
        disc_ids = disc_obj.search(cr, uid, [
            ('state', '=', 'pending'),
            ('date', '<=', today.strftime(
                DEFAULT_SERVER_DATE_FORMAT)),
        ], context=context)

        self.loan_state_done(cr, uid, disc_ids, context=context)

        return True
       
class ng_state_payroll_demise(models.Model):
    '''
    Payroll Employee Demise
    '''
    _name = "ng.state.payroll.demise"
    _description = 'Payroll Employee Demise'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'batch_number': fields.char('Batch Group', help='Batch Group; for batch approval', required=False),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('pending', 'Pending'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
    }
 
    _rec_name = 'date'
     
    @api.multi
    def confirm(self):
        _logger.info("confirm - %s", 'confirm')
        self.write({'state':'confirm'})
     
    @api.multi
    def cancel(self):
        _logger.info("cancel - %s", 'cancel')
        self.write({'state':'cancel'})
    
    @api.multi
    def pending(self):
        _logger.info("pending - %s", 'pending')
        self.write({'state':'pending'})
   
    def _get_default_domain_employees(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        employee_obj = self.pool.get('hr.employee')

        this_user = users_obj.browse(cr, uid, uid, context=context)
        employees = []
        if this_user.domain_mdas:
            employees = employee_obj.search(cr, uid, [('department_id.id', 'in', this_user.domain_mdas.ids)], context=context)
        else:
            employees = employee_obj.search(cr, uid, [], context=context)

        return employees
    
    _defaults = {
        'employee_id': _get_default_domain_employees,
        'state': 'draft',
    }
       
    _track = {
        'state': {
            'ng_state_payroll_demise.mt_alert_demise_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_demise.mt_alert_demise_pending':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'pending',
            'ng_state_payroll_demise.mt_alert_demise_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }

    def _check_state(self, cr, uid, employee_id, effective_date, context=None):
        _logger.info("_check_state - %d", employee_id)
                
        return True
    
    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        _logger.info("_needaction_domain_get - %s", users_obj)

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            _logger.info("_needaction_domain_get - is HR Manager")
            domain = [('state', '=', 'confirm')]
            return domain

        return False

    
    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Demise action!'),
                    _('Demise action has been initiated. Either cancel the demise action or create another to undo it.')
                )

        return super(ng_state_payroll_demise, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for disc in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                disc.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def demise_state_confirm(self, cr, uid, ids, context=None):
        _logger.info("before state_confirm - %d", uid)
        for disc in self.browse(cr, uid, ids, context=context):
            if self._check_state(
                cr, uid, disc.employee_id.id, disc.date, context=context):
                self.write(cr, uid, disc.id, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def demise_state_done(self, cr, uid, ids, context=None):
        _logger.info("Calling demise_state_done...")
        employee_obj = self.pool.get('hr.employee')
        today = datetime.now().date()

        for o in self.browse(cr, uid, ids, context=context):
            if datetime.strptime(
                o.date, DEFAULT_SERVER_DATE_FORMAT
            ).date() <= today and o.state == 'pending':
                if self._check_state(
                    cr, uid, o.employee_id.id, o.date,
                    context=context):
                    status_obj = self.pool.get('ng.state.payroll.status')
                    death_status_ids = status_obj.search(cr, uid, [('name', '=', 'DEATH')], context=context)
                    employee_obj.write(
                        cr, uid, o.employee_id.id, {
                            'status_id': death_status_ids[0]},
                        context=context)
                    self.write(
                        cr, uid, o.id, {'state': 'approved'}, context=context)
                    hrevent_obj = self.pool.get('ng.state.payroll.hrevent')
                    hrevent_obj.create(cr, uid, {'employee_id':o.employee_id.id, 'activity_type':'demise', 'activity_id':o.id})
            else:
                return False
        cr.commit()
        return True

    def try_pending_demise_actions(self, cr, uid, context=None):
        """Completes pending demise actions. Called from
        the scheduler."""

        _logger.info("Running try_pending_demise_actions cron-job...")
        
        disc_obj = self.pool.get('ng.state.payroll.demise')
        today = datetime.now().date()
        disc_ids = disc_obj.search(cr, uid, [
            ('state', '=', 'pending'),
            ('date', '<=', today.strftime(
                DEFAULT_SERVER_DATE_FORMAT)),
        ], context=context)

        self.demise_state_done(cr, uid, disc_ids, context=context)

        return True
   
class ng_state_payroll_termination(models.Model):
    '''
    Payroll Employee Termination
    '''
    _name = "ng.state.payroll.termination"
    _description = 'Payroll Employee Termination'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'batch_number': fields.char('Batch Group', help='Batch Group; for batch approval', required=False),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('pending', 'Pending'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'reason': fields.text('Reason', required=True, help='Reason for termination'),
    }
 
    _rec_name = 'date'
    
    @api.multi
    def confirm(self):
        _logger.info("confirm - %s", 'confirm')
        self.write({'state':'confirm'})
     
    @api.multi
    def cancel(self):
        _logger.info("cancel - %s", 'cancel')
        self.write({'state':'cancel'})
    
    @api.multi
    def pending(self):
        _logger.info("pending - %s", 'pending')
        self.write({'state':'pending'})
   
    def _get_default_domain_employees(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        employee_obj = self.pool.get('hr.employee')

        this_user = users_obj.browse(cr, uid, uid, context=context)
        employees = []
        if this_user.domain_mdas:
            employees = employee_obj.search(cr, uid, [('department_id.id', 'in', this_user.domain_mdas.ids)], context=context)
        else:
            employees = employee_obj.search(cr, uid, [], context=context)

        return employees
    
    _defaults = {
        'employee_id': _get_default_domain_employees,
        'state': 'draft',
    }
       
    _track = {
        'state': {
            'ng_state_payroll_termination.mt_alert_termination_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_termination.mt_alert_termination_pending':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'pending',
            'ng_state_payroll_termination.mt_alert_termination_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }

    def _check_state(self, cr, uid, employee_id, effective_date, context=None):
        _logger.info("_check_state - %d", employee_id)
                
        return True
    
    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        _logger.info("_needaction_domain_get - %s", users_obj)

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            _logger.info("_needaction_domain_get - is HR Manager")
            domain = [('state', '=', 'confirm')]
            return domain

        return False

    
    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Termination action!'),
                    _('Termination action has been initiated. Either cancel the termination action or create another to undo it.')
                )

        return super(ng_state_payroll_termination, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for disc in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                disc.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def termination_state_confirm(self, cr, uid, ids, context=None):
        _logger.info("before state_confirm - %d", uid)
        for disc in self.browse(cr, uid, ids, context=context):
            if self._check_state(
                cr, uid, disc.employee_id.id, disc.date, context=context):
                self.write(cr, uid, disc.id, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def termination_state_done(self, cr, uid, ids, context=None):
        _logger.info("Calling termination_state_done...")
        employee_obj = self.pool.get('hr.employee')
        today = datetime.now().date()

        for o in self.browse(cr, uid, ids, context=context):
            if datetime.strptime(
                o.date, DEFAULT_SERVER_DATE_FORMAT
            ).date() <= today and o.state == 'pending':
                if self._check_state(
                    cr, uid, o.employee_id.id, o.date,
                    context=context):
                    status_obj = self.pool.get('ng.state.payroll.status')
                    death_status_ids = status_obj.search(cr, uid, [('name', '=', 'TERMINATED')], context=context)
                    employee_obj.write(
                        cr, uid, o.employee_id.id, {
                            'status_id': death_status_ids[0]},
                        context=context)
                    self.write(
                        cr, uid, o.id, {'state': 'approved'}, context=context)
                    hrevent_obj = self.pool.get('ng.state.payroll.hrevent')
                    hrevent_obj.create(cr, uid, {'employee_id':o.employee_id.id, 'activity_type':'termination', 'activity_id':o.id})
            else:
                return False
        cr.commit()
        return True

    def try_pending_termination_actions(self, cr, uid, context=None):
        """Completes pending termination actions. Called from
        the scheduler."""

        _logger.info("Running try_pending_termination_actions cron-job...")
        
        disc_obj = self.pool.get('ng.state.payroll.termination')
        today = datetime.now().date()
        disc_ids = disc_obj.search(cr, uid, [
            ('state', '=', 'pending'),
            ('date', '<=', today.strftime(
                DEFAULT_SERVER_DATE_FORMAT)),
        ], context=context)

        self.termination_state_done(cr, uid, disc_ids, context=context)

        return True
   
class ng_state_payroll_query(models.Model):
    '''
    HR Employee Query
    '''
    _name = "ng.state.payroll.query"
    _description = 'HR Employee Query'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'batch_number': fields.char('Batch Group', help='Batch Group; for batch approval', required=False),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('pending', 'Pending'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'title': fields.char('Title', required=True, help='Title of the query'),
        'comments': fields.text('Comments', required=True, help='Description of the query'),
        'emp_response': fields.text('Response', help='Response to query'),
    }
 
    _rec_name = 'date'
     
    @api.multi
    def confirm(self):
        _logger.info("confirm - %s", 'confirm')
        self.write({'state':'confirm'})
     
    @api.multi
    def cancel(self):
        _logger.info("cancel - %s", 'cancel')
        self.write({'state':'cancel'})
    
    @api.multi
    def pending(self):
        _logger.info("pending - %s", 'pending')
        self.write({'state':'pending'})
  
    def _get_default_domain_employees(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        employee_obj = self.pool.get('hr.employee')

        this_user = users_obj.browse(cr, uid, uid, context=context)
        employees = []
        if this_user.domain_mdas:
            employees = employee_obj.search(cr, uid, [('department_id.id', 'in', this_user.domain_mdas.ids)], context=context)
        else:
            employees = employee_obj.search(cr, uid, [], context=context)

        return employees
    
    _defaults = {
        'employee_id': _get_default_domain_employees,
        'state': 'draft',
    }
       
    _track = {
        'state': {
            'ng_state_payroll_query.mt_alert_query_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_query.mt_alert_query_pending':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'pending',
            'ng_state_payroll_query.mt_alert_query_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }

    def _check_state(self, cr, uid, employee_id, effective_date, context=None):
        _logger.info("_check_state - %d", employee_id)
                
        return True
    
    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        _logger.info("_needaction_domain_get - %s", users_obj)

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            _logger.info("_needaction_domain_get - is HR Manager")
            domain = [('state', '=', 'confirm')]
            return domain

        return False

    
    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete query action!'),
                    _('query action has been initiated. Either cancel the query action or create another to undo it.')
                )

        return super(ng_state_payroll_query, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for disc in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                disc.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def query_state_confirm(self, cr, uid, ids, context=None):
        _logger.info("before state_confirm - %d", uid)
        for qry in self.browse(cr, uid, ids, context=context):
            if self._check_state(
                cr, uid, qry.employee_id.id, qry.date, context=context):
                self.write(cr, uid, qry.id, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def query_state_done(self, cr, uid, ids, context=None):
        today = datetime.now().date()
        for o in self.browse(cr, uid, ids, context=context):
            if datetime.strptime(
                o.date, DEFAULT_SERVER_DATE_FORMAT
            ).date() <= today and o.state == 'pending':
                if self._check_state(
                    cr, uid, o.employee_id.id, o.date,
                    context=context):
                    #If third approved query, initiate termination process
                    self.write(cr, uid, o.id, {'state': 'approved'}, context=context)
                    hrevent_obj = self.pool.get('ng.state.payroll.hrevent')
                    hrevent_obj.create(cr, uid, {'employee_id':o.employee_id.id, 'activity_type':'demise', 'activity_id':o.id})
                    query_obj = self.pool.get('ng.state.payroll.query')
                    query_ids = query_obj.search(cr, uid, [('employee_id', '=', o.employee_id.id),('state', '=', 'approved')], context=context)
                    if len(query_ids) == 3:
                        #Create termination request workflow
                        termination_obj = self.pool.get('ng.state.payroll.termination')
                        termination_obj.create(cr, uid, {'employee_id':o.employee_id.id,'state':'draft','date':today.strftime(DEFAULT_SERVER_DATE_FORMAT),'reason':'Third query issued.'}, context=context)
            else:
                return False
        cr.commit()
        return True

    def try_pending_query_actions(self, cr, uid, context=None):
        """Completes pending query actions. Called from
        the scheduler."""
        
        _logger.info("Running try_pending_query_actions cron-job...")

        query_obj = self.pool.get('ng.state.payroll.query')
        today = datetime.now().date()
        query_ids = query_obj.search(cr, uid, [
            ('state', '=', 'pending'),
            ('date', '<=', today.strftime(
                DEFAULT_SERVER_DATE_FORMAT)),
        ], context=context)

        self.query_state_done(cr, uid, query_ids, context=context)

        return True
   
class ng_state_payroll_retirement(models.Model):
    '''
    Payroll Employee Retirement
    '''
    _name = "ng.state.payroll.retirement"
    _description = 'Payroll Employee Retirement'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'batch_number': fields.char('Batch Group', help='Batch Group; for batch approval', required=False),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('pending', 'Pending'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'retirement_type': fields.selection([
            ('auto', 'Automatic'),
            ('voluntary', 'Voluntary'),
        ], 'Type', required=True, readonly=True),           
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'payscheme_id': fields.related('employee_id', 'payscheme_id', type='many2one', relation='ng.state.payroll.payscheme', string='Pay Scheme', readonly=1),
        'department_id': fields.related('employee_id', 'department_id', type='many2one', relation='hr.department', string='Organization', readonly=1),
        'birthday': fields.related('employee_id', 'birthday', type='date', string='Birth', readonly=1),
        'hire_date': fields.related('employee_id', 'hire_date', type='date', string='First Hire', readonly=1),
        'retirement_due_date': fields.related('employee_id', 'retirement_due_date', type='date', string='Due Date', readonly=1),
        'retirement_index': fields.related('employee_id', 'retirement_index', type='char', string='Index', readonly=1),
    }
 
    _rec_name = 'date'
    
    @api.multi
    def confirm(self):
        _logger.info("confirm - %s", 'confirm')
        self.write({'state':'confirm'})
     
    @api.multi
    def cancel(self):
        _logger.info("cancel - %s", 'cancel')
        self.write({'state':'cancel'})
    
    @api.multi
    def pending(self):
        _logger.info("pending - %s", 'pending')
        self.write({'state':'pending'})
   
    def _get_default_domain_employees(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        employee_obj = self.pool.get('hr.employee')

        this_user = users_obj.browse(cr, uid, uid, context=context)
        employees = []
        if this_user.domain_mdas:
            employees = employee_obj.search(cr, uid, [('department_id.id', 'in', this_user.domain_mdas.ids)], context=context)
        else:
            employees = employee_obj.search(cr, uid, [], context=context)

        return employees
    
    _defaults = {
        'employee_id': _get_default_domain_employees,
        'state': 'draft',
        'retirement_type': 'voluntary',
    }
       
    _track = {
        'state': {
            'ng_state_payroll_retirement.mt_alert_retirement_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_retirement.mt_alert_retirement_pending':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'pending',
            'ng_state_payroll_retirement.mt_alert_retirement_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }

    def onchange_employee(self, cr, uid, ids, employee_id, context=None):
        res = {'value': {'employee_id': employee_id}}

        if employee_id:
            ee = self.pool.get('hr.employee').browse(cr, uid, employee_id, context=context)
            res['value']['payscheme_id'] = ee.payscheme_id.id
            res['value']['department_id'] = ee.department_id.id
            res['value']['birthday'] = ee.birthday
            res['value']['hire_date'] = ee.hire_date
            res['value']['retirement_due_date'] = ee.retirement_due_date
            res['value']['retirement_index'] = ee.retirement_index
        return res
    
    def _check_state(self, cr, uid, employee_id, effective_date, context=None):
        _logger.info("_check_state - %d", employee_id)
                
        return True
    
    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        _logger.info("_needaction_domain_get - %d", uid)

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            _logger.info("_needaction_domain_get - is HR Manager")
            domain = [('state', '=', 'confirm')]
            return domain

        return False
    
    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Promotion!'),
                    _('Retirement process has been initiated. Either cancel the retirement process or create another to undo it.')
                )

        return super(ng_state_payroll_retirement, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for disc in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                disc.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def retirement_state_confirm(self, cr, uid, ids, context=None):
        _logger.info("before state_confirm - %d", uid)
        for disc in self.browse(cr, uid, ids, context=context):
            if self._check_state(
                cr, uid, disc.employee_id.id, disc.date, context=context):
                self.write(cr, uid, disc.id, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def retirement_state_done(self, cr, uid, ids, context=None):

        employee_obj = self.pool.get('hr.employee')
        today = datetime.now().date()

        for o in self.browse(cr, uid, ids, context=context):
            if datetime.strptime(
                o.date, DEFAULT_SERVER_DATE_FORMAT
            ).date() <= today and o.state == 'pending':
                if self._check_state(
                    cr, uid, o.employee_id.id, o.date,
                    context=context):
                    status_obj = self.pool.get('ng.state.payroll.status')
                    retirement_status_ids = status_obj.search(cr, uid, [('name', '=', 'RETIRED')], context=context)
                    employee_obj.write(
                        cr, uid, o.employee_id.id, {
                            'status_id': retirement_status_ids[0]},
                        context=context)
                    self.write(
                        cr, uid, o.id, {'state': 'approved'}, context=context)
                    hrevent_obj = self.pool.get('ng.state.payroll.hrevent')
                    hrevent_obj.create(cr, uid, {'employee_id':o.employee_id.id, 'activity_type':'retirement', 'activity_id':o.id})
            else:
                return False
        cr.commit()
        return True
    
    def try_init_due_retirements(self, cr, uid, context=None):
        _logger.info("Running try_init_due_retirements cron-job...")
        employee_obj = self.pool.get('hr.employee')
        employee_ids = employee_obj.search(cr, uid, [('active', '=', True), ('retirement_due_date', '=', False), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id', context=context)
        _logger.info("try_init_due_retirements - employees=%d", len(employee_ids))

        for emp in employee_obj.browse(cr, uid, employee_ids, context=context):                          
            #Use hire date and date of birth to calculate retirement date
            retirement_date = False
            retirement_date_dofa = False
            retirement_date_dob = False
            retirement_index = False
            if emp.payscheme_id.use_dofa:
                retirement_date_dofa = datetime.strptime(emp.hire_date, DEFAULT_SERVER_DATE_FORMAT) + relativedelta(years=emp.payscheme_id.service_years)
                retirement_date = retirement_date_dofa
                retirement_index = 'dofa'
            if emp.payscheme_id.use_dob:
                retirement_date_dob = datetime.strptime(emp.birthday, DEFAULT_SERVER_DATE_FORMAT) + relativedelta(years=emp.payscheme_id.retirement_age)
                retirement_date = retirement_date_dob
                retirement_index = 'dofb'
            if emp.payscheme_id.use_dofa and emp.payscheme_id.use_dob:
                if retirement_date_dofa < retirement_date_dob:
                    retirement_date = retirement_date_dofa
                    retirement_index = 'dofa'
                else:
                    retirement_date = retirement_date_dob
                    retirement_index = 'dofb'
            if retirement_date:
                cr.execute("update hr_employee set retirement_due_date='" + retirement_date.strftime(DEFAULT_SERVER_DATE_FORMAT) + "',retirement_index='" + retirement_index + "' where id=" + str(emp.id))
            else:
                cr.execute("update hr_employee set retirement_due_date='3000-01-01' where id=" + str(emp.id))
        cr.commit()
        return True
    
    def try_due_retirements(self, cr, uid, context=None):
        _logger.info("Running try_due_retirements cron-job...")
        today = datetime.now().date()
        employee_obj = self.pool.get('hr.employee')
        retirement_obj = self.pool.get('ng.state.payroll.retirement')
        cr.execute("select employee_id from ng_state_payroll_retirement")
        emp_retire_req_ids = cr.fetchall()
        
        employee_ids = employee_obj.search(cr, uid, [('id', 'not in', emp_retire_req_ids), ('active', '=', True), ('retirement_due_date', '!=', False), ('retirement_due_date', '<=', today.strftime(DEFAULT_SERVER_DATE_FORMAT)), '|', ('status_id.name', '=', 'ACTIVE'), ('status_id.name', '=', 'SUSPENDED')], order='id', context=context)
        _logger.info("try_due_retirements - employees=%d", len(employee_ids))

        for emp in employee_obj.browse(cr, uid, employee_ids, context=context):                          
            retirement_id = retirement_obj.create(cr, uid, {
                'employee_id':emp.id,
                'retirement_type':'auto',
                'date':today.strftime(DEFAULT_SERVER_DATE_FORMAT),
            }, context=context)
            self.retirement_state_confirm(cr, uid, retirement_id, context=context)
        cr.commit()
        return True
                
    def try_pending_retirement_actions(self, cr, uid, context=None):
        """Completes pending retirement actions. Called from
        the scheduler."""
        
        _logger.info("Running try_pending_retirement_actions cron-job...")

        retirement_obj = self.pool.get('ng.state.payroll.retirement')
        today = datetime.now().date()
        retirement_ids = retirement_obj.search(cr, uid, [
            ('state', '=', 'pending'),
            ('date', '<=', today.strftime(
                DEFAULT_SERVER_DATE_FORMAT)),
        ], context=context)

        self.retirement_state_done(cr, uid, retirement_ids, context=context)
        return True
        
class hr_transfer(orm.Model):

    _name = 'hr.department.transfer'
    _description = 'MDA Transfer'

    _inherit = ['mail.thread', 'ir.needaction_mixin']

    _columns = {
        'batch_number': fields.char('Batch Group', help='Batch Group; for batch approval', required=False),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'src_department_id': fields.many2one('hr.department', 'From MDA', required=True),
        'dst_department_id': fields.many2one('hr.department', 'To MDA', required=True),
        'src_school_id': fields.many2one('ng.state.payroll.school', 'From School', required=False),
        'dst_school_id': fields.many2one('ng.state.payroll.school', 'To School', required=False),
        'error_msg': fields.char('Error Message', help='Error Message holding up process', required=False),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('pending', 'Pending'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ],
            'State', readonly=True),
    }

    _rec_name = 'date'
    
    @api.multi
    def confirm(self):
        _logger.info("confirm - %s", 'confirm')
        self.write({'state':'confirm'})
     
    @api.multi
    def cancel(self):
        _logger.info("cancel - %s", 'cancel')
        self.write({'state':'cancel'})
    
    @api.multi
    def pending(self):
        _logger.info("pending - %s", 'pending')
        self.write({'state':'pending'})
        
    @api.onchange('src_department_id')
    def src_school_id_update(self):
        return {'domain': {'src_school_id': [('org_id','=',self.src_department_id.id)] }}

    @api.onchange('dst_department_id')
    def dst_school_id_update(self):
        return {'domain': {'dst_school_id': [('org_id','=',self.dst_department_id.id)] }}
   
    def _get_default_domain_employees(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        employee_obj = self.pool.get('hr.employee')

        this_user = users_obj.browse(cr, uid, uid, context=context)
        employees = []
        if this_user.domain_mdas:
            employees = employee_obj.search(cr, uid, [('department_id.id', 'in', this_user.domain_mdas.ids)], context=context)
        else:
            employees = employee_obj.search(cr, uid, [], context=context)

        return employees
    
    _defaults = {
        'employee_id': _get_default_domain_employees,
        'state': 'draft',
    }

    _track = {
        'state': {
            'hr_transfer.mt_alert_item_obj_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'hr_transfer.mt_alert_item_obj_pending':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'pending',
            'hr_transfer.mt_alert_item_obj_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }

    def _needaction_domain_get(self, cr, uid, context=None):

        users_obj = self.pool.get('res.users')

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            domain = [('state', '=', 'confirm')]
            return domain

        return False

    
    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Promotion!'),
                    _('Promotion has been initiated. Either cancel the promotion or create another promotion to undo it.')
                )

        return super(hr_transfer, self).unlink(cr, uid, ids, context=context)

    def onchange_employee(self, cr, uid, ids, employee_id, context=None):

        res = {'value': {'src_department_id': False}}

        if employee_id:
            ee = self.pool.get('hr.employee').browse(
                cr, uid, employee_id, context=context)
            res['value']['src_department_id'] = ee.department_id.id
            if ee.school_id:
                res['value']['src_school_id'] = ee.school_id.id

        return res

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for item_obj in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                item_obj.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def _check_state(self, cr, uid, employee_id, effective_date, context=None):
        _logger.info("_check_state - %d", employee_id)
        employee_obj = self.pool.get('hr.employee')
        data = employee_obj.read(
            cr, uid, employee_id, ['state', 'retirement_due_date'], context=context) 
        if data.get('retirement_due_date', False) and data['retirement_due_date'] != '':
            retirementDate = datetime.strptime(
                data['retirement_due_date'], DEFAULT_SERVER_DATE_FORMAT)
            dEffective = datetime.strptime(
                effective_date, DEFAULT_SERVER_DATE_FORMAT)
            if dEffective >= retirementDate:
                return False
                
        return True

    def state_confirm(self, cr, uid, ids, context=None):
        _logger.info("before state_confirm - %d", uid)
        for item_obj in self.browse(cr, uid, ids, context=context):
            if self._check_state(
                cr, uid, item_obj.employee_id.id, item_obj.date, context=context):
                self.write(cr, uid, item_obj.id, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def state_done(self, cr, uid, ids, context=None):

        employee_obj = self.pool.get('hr.employee')
        today = datetime.now().date()

        for item_obj in self.browse(cr, uid, ids, context=context):
            if datetime.strptime(
                item_obj.date, DEFAULT_SERVER_DATE_FORMAT
            ).date() <= today and item_obj.state == 'pending':
                #Add school transfer
                transfer_dict = {'department_id': item_obj.dst_department_id.id}
                if item_obj.dst_school_id:
                    transfer_dict.update({'school_id': item_obj.dst_school_id.id})
                employee_obj.write(
                    cr, uid, item_obj.employee_id.id, transfer_dict,
                    context=context)
                self.write(
                    cr, uid, item_obj.id, {'state': 'approved'}, context=context)
                hrevent_obj = self.pool.get('ng.state.payroll.hrevent')
                hrevent_obj.create(cr, uid, {'employee_id':item_obj.employee_id.id, 'activity_type':'transfer', 'activity_id':item_obj.id})
            else:
                return False
        cr.commit()
        return True

    def try_pending_department_transfers(self, cr, uid, context=None):
        """Completes pending departmental transfers. Called from
        the scheduler."""

        _logger.info("Running try_pending_department_transfers cron-job...")
        
        item_singleton = self.pool.get('hr.department.transfer')
        today = datetime.now().date()
        item_obj_ids = item_singleton.search(cr, uid, [
            ('state', '=', 'pending'),
            ('date', '<=', today.strftime(
                DEFAULT_SERVER_DATE_FORMAT)),
        ], context=context)

        self.state_done(cr, uid, item_obj_ids, context=context)

        return True
   
class ng_state_payroll_changereq(models.Model):
    '''
    Payroll Employee Change Request
    '''
    _name = "ng.state.payroll.changereq"
    _description = 'Payroll Employee Change Request'
    _inherit = ['mail.thread', 'ir.needaction_mixin']
    
    _columns = {
        'batch_number': fields.char('Batch Group', help='Batch Group; for batch approval', required=False),
        'employee_id': fields.many2one('hr.employee', 'Employee', ondelete='cascade', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('confirm', 'Confirmed'),
            ('pending', 'Pending'),
            ('approved', 'Approved'),
            ('cancel', 'Cancelled'),
        ], 'State', readonly=True),
        'active_flag': fields.boolean('Active', help='Active Status', required=True),
        'date': fields.date('Effective Date', required=True, readonly=True, states={'draft': [('readonly', False)]}),
        'name_related': fields.char('Employee Name', help='Employee Name'),
        'sinid': fields.char('Pension PIN', help='Pension PIN'),
        'ssnid': fields.char('SSN', help='Social Security Number'),
        'employee_no': fields.char('Employee Number', help='Employee Number'),
        'school_emp_id': fields.char('School Employee ID', help='School Employee ID', required=False),
        'bank_account_no': fields.char('Bank Account', help='Bank Account Number'),
        'hire_date': fields.date('Hire Date', help='Date of Hire'),
        'confirmation_date': fields.date('Confirmation Date', help='Date of Confirmation'),
        'lga_id': fields.many2one('ng.state.payroll.lga', 'LGA'),
        'pfa_id': fields.many2one('ng.state.payroll.pfa', 'PFA'),
        'school_id': fields.many2one('ng.state.payroll.school', 'School', required=False),
        'paycategory_id': fields.many2one('ng.state.payroll.paycategory', 'Pay Category'),
        'payscheme_id': fields.many2one('ng.state.payroll.payscheme', 'Pay Scheme'),
        'level_id': fields.many2one('ng.state.payroll.level', 'Grade'),
        'grade_level': fields.selection([
            (1, 'GL-1'),
            (2, 'GL-2'),
            (3, 'GL-3'),
            (4, 'GL-4'),
            (5, 'GL-5'),
            (6, 'GL-6'),
            (7, 'GL-7'),
            (8, 'GL-8'),
            (9, 'GL-9'),
            (10, 'GL-10'),
            (12, 'GL-12'),
            (13, 'GL-13'),
            (14, 'GL-14'),
            (15, 'GL-15'),
            (16, 'GL-16'),
            (17, 'GL-17'),
            (18, 'GL-18'),
            (19, 'GL-19'),
            (20, 'GL-20'),
        ], 'Grade Level', readonly=True),
        'grade_step': fields.selection([
            (1, 'Step-1'),
            (2, 'Step-2'),
            (3, 'Step-3'),
            (4, 'Step-4'),
            (5, 'Step-5'),
            (6, 'Step-6'),
            (7, 'Step-7'),
            (8, 'Step-8'),
            (9, 'Step-9'),
            (10, 'Step-10'),
            (11, 'Step-11'),
            (12, 'Step-12'),
            (13, 'Step-13'),
            (14, 'Step-14'),
            (15, 'Step-15'),
            (16, 'Step-16'),
            (17, 'Step-17'),
            (18, 'Step-18'),
            (19, 'Step-19'),
            (20, 'Step-20'),
        ], 'Grade Step'),
        'title_id': fields.many2one('res.partner.title', 'Title'),
        'status_id': fields.many2one('ng.state.payroll.status', 'Employee Status'),
        'bank_id': fields.many2one('res.bank', string='Bank'),
        'pensiontype_id': fields.many2one('ng.state.payroll.pensiontype', 'Pension Type', required=False),
        'tco_id': fields.many2one('ng.state.payroll.tco', 'TCO', required=False),
        'pensionfile_no': fields.char('Pension File', help='Pension File Number'),
        'annual_pension': fields.float('Annual Pension', help='Annual Pension'),
    }
 
    _rec_name = 'date'
    
    _defaults = {
        'state': 'draft',
    }

    _track = {
        'state': {
            'ng_state_payroll_changereq.mt_alert_changereq_confirmed':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'confirm',
            'ng_state_payroll_changereq.mt_alert_changereq_pending':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'pending',
            'ng_state_payroll_changereq.mt_alert_changereq_done':
                lambda self, cr, uid, obj, ctx=None: obj['state'] == 'approved',
        },
    }
    
    @api.multi
    def confirm(self):
        _logger.info("confirm - %s", 'confirm')
        self.write({'state':'confirm'})
     
    @api.multi
    def cancel(self):
        _logger.info("cancel - %s", 'cancel')
        self.write({'state':'cancel'})
    
    @api.multi
    def pending(self):
        _logger.info("pending - %s", 'pending')
        self.write({'state':'pending'})
        
    @api.onchange('department_id')
    def school_id_update(self):
        return {'domain': {'school_id': [('org_id','=',self.department_id.id)] }}
        
    @api.onchange('payscheme_id')
    def level_id_update(self):
        return {'domain': {'level_id': [('paygrade_id.payscheme_id','=',self.payscheme_id.id)] }}
    
    def onchange_employee(self, cr, uid, ids, employee_id, context=None):
        _logger.info("onchange_employee - %s", self.state)
        res = {'value': {'employee_id': employee_id}}

        if employee_id and self.state=='draft':
            ee = self.pool.get('hr.employee').browse(cr, uid, employee_id, context=context)
            res['value']['name_related'] = ee.name_related
            res['value']['employee_no'] = ee.employee_no
            res['value']['school_emp_id'] = ee.school_emp_id
            res['value']['bank_account_no'] = ee.bank_account_no
            res['value']['hire_date'] = ee.hire_date
            res['value']['pensionfile_no'] = ee.pensionfile_no
            res['value']['annual_pension'] = ee.annual_pension
            res['value']['active_flag'] = ee.active
            if ee.confirmation_date:
                res['value']['confirmation_date'] = ee.confirmation_date
            if ee.lga_id:
                res['value']['lga_id'] = ee.lga_id.id
            if ee.pfa_id:
                res['value']['pfa_id'] = ee.pfa_id.id
            if ee.school_id:
                res['value']['school_id'] = ee.school_id.id
            if ee.paycategory_id:
                res['value']['paycategory_id'] = ee.paycategory_id.id
            if ee.payscheme_id:
                res['value']['payscheme_id'] = ee.payscheme_id.id
            if ee.level_id:
                res['value']['level_id'] = ee.level_id.id
                res['value']['grade_step'] = ee.level_id.step
            if ee.title_id:
                res['value']['title_id'] = ee.title_id.id
            if ee.status_id:
                res['value']['status_id'] = ee.status_id.id
            if ee.bank_id:
                res['value']['bank_id'] = ee.bank_id.id
            if ee.pensiontype_id:
                res['value']['pensiontype_id'] = ee.pensiontype_id.id
            if ee.tco_id:
                res['value']['tco_id'] = ee.tco_id.id
        return res
    
    def _check_state(self, cr, uid, employee_id, effective_date, context=None):
        _logger.info("_check_state - %d", employee_id)
                
        return True
    
    def _needaction_domain_get(self, cr, uid, context=None):
        users_obj = self.pool.get('res.users')
        _logger.info("_needaction_domain_get - %s", users_obj)

        if users_obj.has_group(cr, uid, 'base.group_hr_manager'):
            _logger.info("_needaction_domain_get - is HR Manager")
            domain = [('state', '=', 'confirm')]
            return domain

        return False
    
    def unlink(self, cr, uid, ids, context=None):
        for item_obj in self.browse(cr, uid, ids, context=context):
            if item_obj.state not in ['draft']:
                raise osv.except_osv(
                    _('Unable to Delete Change Request action!'),
                    _('Change Request action has been initiated. Either cancel the Change Request action or create another to undo it.')
                )

        return super(ng_state_payroll_changereq, self).unlink(cr, uid, ids, context=context)

    def effective_date_in_future(self, cr, uid, ids, context=None):

        today = datetime.now().date()
        for disc in self.browse(cr, uid, ids, context=context):
            effective_date = datetime.strptime(
                disc.date, DEFAULT_SERVER_DATE_FORMAT).date()
            if effective_date <= today:
                return False

        return True

    def changereq_state_confirm(self, cr, uid, ids, context=None):
        _logger.info("before state_confirm - %d", uid)
        for o in self.browse(cr, uid, ids, context=context):
            if self._check_state(
                cr, uid, o.employee_id.id, o.date, context=context):
                self.write(cr, uid, o.id, {'state': 'confirm'}, context=context)
        _logger.info("after state_confirm - %d", uid)
        cr.commit()
        return True

    def changereq_state_done(self, cr, uid, ids, context=None):
        _logger.info("Calling changereq_state_done...")
        employee_obj = self.pool.get('hr.employee')
        cron_obj = self.pool.get('ir.cron')
        today = datetime.now().date()
        
        resolve_earn_dedt = False
        for o in self.browse(cr, uid, ids, context=context):
            if datetime.strptime(
                o.date, DEFAULT_SERVER_DATE_FORMAT
            ).date() <= today and o.state == 'pending':
                if self._check_state(cr, uid, o.employee_id.id, o.date, context=context):
                    emp_dict = {}
                    emp_dict.update({'active':o.active_flag})
                    if o.employee_no:
                        emp_dict.update({'employee_no':o.employee_no})
                    if o.name_related:
                        emp_dict.update({'name_related':o.name_related})
                    if o.school_emp_id:
                        emp_dict.update({'school_emp_id':o.school_emp_id})
                    if o.bank_account_no:
                        emp_dict.update({'bank_account_no':o.bank_account_no})
                    if o.hire_date:
                        emp_dict.update({'hire_date':o.hire_date})
            
                        retirement_date = False
                        retirement_date_dofa = False
                        retirement_date_dob = False
                        retirement_index = False
                        if o.employee_id.payscheme_id.use_dofa:
                            retirement_date_dofa = datetime.strptime(o.employee_id.hire_date, DEFAULT_SERVER_DATE_FORMAT) + relativedelta(years=o.employee_id.payscheme_id.service_years)
                            retirement_date = retirement_date_dofa
                            retirement_index = 'dofa'
                        if o.employee_id.payscheme_id.use_dob:
                            retirement_date_dob = datetime.strptime(o.employee_id.birthday, DEFAULT_SERVER_DATE_FORMAT) + relativedelta(years=o.employee_id.payscheme_id.retirement_age)
                        retirement_date = retirement_date_dob
                        retirement_index = 'dofb'
                        if o.employee_id.payscheme_id.use_dofa and o.employee_id.payscheme_id.use_dob:
                        if retirement_date_dofa < retirement_date_dob:
                            retirement_date = retirement_date_dofa
                            retirement_index = 'dofa'
                        else:
                            retirement_date = retirement_date_dob
                            retirement_index = 'dofb'
                        if retirement_date:
                            emp_dict.update({'retirement_due_date':retirement_date.strftime('%Y-%m-%d')})
                            emp_dict.update({'retirement_index':retirement_index})                    
                        _logger.info("Name: " + o.employee_id.name_related + "...New Retirement Date: " + retirement_date.strftime('%Y-%m-%d'))
                    if o.confirmation_date:
                        emp_dict.update({'confirmation_date':o.confirmation_date})
                    if o.lga_id:
                        emp_dict.update({'lga_id':o.lga_id.id})
                    if o.pfa_id:
                        emp_dict.update({'pfa_id':o.pfa_id.id})
                    if o.school_id:
                        emp_dict.update({'school_id':o.school_id.id})
                    if o.paycategory_id:
                        emp_dict.update({'paycategory_id':o.paycategory_id.id})
                    if o.payscheme_id:
                        emp_dict.update({'payscheme_id':o.payscheme_id.id})
                        emp_dict.update({'resolved_earn_dedt': False})
                        resolve_earn_dedt = True
                    if o.level_id:
                        emp_dict.update({'level_id':o.level_id.id})
                        emp_dict.update({'resolved_earn_dedt': False})
                        resolve_earn_dedt = True
                    if o.grade_level:
                        emp_dict.update({'grade_level':o.grade_level})
                        emp_dict.update({'resolved_earn_dedt': False})
                        resolve_earn_dedt = True
                            
                    if o.title_id:
                        emp_dict.update({'title_id':o.title_id.id})
                    if o.status_id:
                        emp_dict.update({'status_id':o.status_id.id})
                    if o.bank_id:
                        emp_dict.update({'bank_id':o.bank_id.id})
                    if o.pensiontype_id:
                        emp_dict.update({'pensiontype_id':o.pensiontype_id.id})
                    if o.tco_id:
                        emp_dict.update({'tco_id':o.tco_id.id})
                    if o.pensionfile_no:
                        emp_dict.update({'pensionfile_no':o.pensionfile_no})
                    if o.sinid:
                        emp_dict.update({'sinid':o.sinid})
                    if o.ssnid:
                        emp_dict.update({'ssnid':o.ssnid})
                    if o.annual_pension:                
                        emp_dict.update({'annual_pension':o.annual_pension})
        
                    employee_obj.write(cr, uid, o.employee_id.id, emp_dict, context=context)
                    self.write(cr, uid, o.id, {'state': 'approved'}, context=context)
                    hrevent_obj = self.pool.get('ng.state.payroll.hrevent')
                    hrevent_obj.create(cr, uid, {'employee_id':o.employee_id.id, 'activity_type':'changereq', 'activity_id':o.id})
            else:
                return False

        if resolve_earn_dedt:
            cron_ids = cron_obj.search(cr, uid, [('name', '=', 'Resolve Standard Earnings and Deductions')], context=context)
            cron_rec = cron_obj.browse(cr, uid, cron_ids[0], context=context)
            nextcall = datetime.now() + timedelta(seconds=3)
            cron_rec.write({'nextcall':nextcall.strftime(DEFAULT_SERVER_DATETIME_FORMAT)})
        cr.commit()
        return True

    def try_pending_changereq_actions(self, cr, uid, context=None):
        """Completes pending changereq actions. Called from
        the scheduler."""

        _logger.info("Running try_pending_changereq_actions cron-job...")
        
        disc_obj = self.pool.get('ng.state.payroll.changereq')
        today = datetime.now().date()
        disc_ids = disc_obj.search(cr, uid, [
            ('state', '=', 'pending'),
            ('date', '<=', today.strftime(
                DEFAULT_SERVER_DATE_FORMAT)),
        ], context=context)

        self.changereq_state_done(cr, uid, disc_ids, context=context)

        return True

class ng_state_payroll_batchapproval(models.Model):
    '''
    Batch Approval
    '''
    _name = "ng.state.payroll.batchapproval"
    _description = 'Batch approval of employee related actions'

    _columns = {
        'name': fields.char('Batch Name', help='Batch Name', required=True),
        'batch_number': fields.char('Batch Group', help='Batch Group; for batch approval', required=True),
        'start_date': fields.date('Start Date', help='Start Date', required=False),
        'end_date': fields.date('End Date', help='End Date', required=False),
        'state': fields.selection([
            ('draft', 'Draft'),
            ('processed', 'Complete'),
        ], 'State', required=True, readonly=True),
        'action_type': fields.selection([
            ('ng_state_payroll_disciplinary', 'Disciplinary'),
            ('ng_state_payroll_changereq', 'Change Request'),
            ('ng_state_payroll_retirement', 'Retirement'),
            ('ng_state_payroll_query', 'Query'),
            ('ng_state_payroll_termination', 'Termination'),
            ('ng_state_payroll_demise', 'Demise'),
            ('ng_state_payroll_promotion', 'Promotion'),
            ('ng_state_payroll_loan', 'Loan'),
            ('hr_transfer', 'HR Transfer'),
        ], 'HR Action', required=True),
    }
    
    _defaults = {
        'state': 'draft',
    }

    @api.multi        
    def process(self, context=None):
        sql_string = "update " + self.action_type + " set state='pending' where batch_number='" + self.batch_number + "'"
        if self.start_date and self.end_date:
            sql_string = sql_string + " and date >= '" + datetime.strptime(self.start_date, '%Y-%m-%d').strftime('%Y-%m-%d') + "' and date <= '" + datetime.strptime(self.end_date, '%Y-%m-%d').strftime('%Y-%m-%d') + "'"
        self.env.cr.execute(sql_string)
        self.write({'state': 'processed'})    
