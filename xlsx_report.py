import logging
import string
import os

from openerp.osv import osv
from openerp.addons.report_xlsx.report.report_xlsx import ReportXlsx

_logger = logging.getLogger(__name__)

SUPERUSER_ID = 1
REPORTS_DIR = '/odoo/odoo9/reports/'

if not os.path.exists(REPORTS_DIR):
    os.makedirs(REPORTS_DIR)

class payroll_summary_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            _logger.info("Creating report...")

            sheet = workbook.add_worksheet(payroll_objs[0].name[:31])
            row = 0
            indices = [0,1,2,3,4,5,6]
            header = ['Serial #','MDA','Gross Income','Taxable Income','Net Income','PAYE Tax','Leave Allowance']
            for c in indices:
                sheet.write(row, c, header[c], bold_font)
            
            row = 1
            summary_items = payroll_objs[0].payroll_summary_ids    
            for summary_item in summary_items:
                sheet.write_number(row, 0, row)
                if summary_item.department_id:
                    sheet.write_string(row, 1, summary_item.department_id.name)
                else:
                    sheet.write_string(row, 1, '')
                sheet.write_number(row, 2, summary_item.total_gross_income, money_format)
                sheet.write_number(row, 3, summary_item.total_taxable_income, money_format)
                sheet.write_number(row, 4, summary_item.total_net_income, money_format)
                sheet.write_number(row, 5, summary_item.total_paye_tax, money_format)
                sheet.write_number(row, 6, summary_item.total_leave_allowance, money_format)
                row += 1
                
            #Sum up
            sheet.write_string(row, 0, 'TOTAL', bold_font)
            for col in [1,2,3,4,5,6]:
                col_name = string.ascii_uppercase[col]
                sheet.write_formula(row, col, '=SUM(' + col_name + '2:' + col_name + str(row) + ')', money_format)

            workbook.close()
            xlsx_data = output.getvalue()
            #payroll_objs[0].update({'summary_report': xlsx_data})
        return xlsx_data

class pension_exec_summary_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        header_font = workbook.add_format({'font_size': 18, 'bold': True, 'bottom': 2, 'top': 2})
        header_font2 = workbook.add_format({'font_size': 14, 'bold': True, 'align': 'center', 'bottom': 2, 'top': 2})
        header_font3 = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2})
        header_font3_money_format = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2, 'num_format': '###,###,##0.#0'})
        header_font3_int_format = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2, 'num_format': '###,###,###'})
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        money_format_bold = workbook.add_format({'num_format': '###,###,##0.#0','bold': True})
        format11 = workbook.add_format()
        format11.set_num_format('000')
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            _logger.info("Creating report...")

            sheet = workbook.add_worksheet(payroll_objs[0].name[:31])
            sheet.insert_image('A1:H5', '/odoo/odoo9/osun_ippms.png')
            sheet.merge_range('A6:H7', 'Osun State Executive Pension Summary for: ' + payroll_objs[0].name + ' (' + payroll_objs[0].calendar_id.name + ')', header_font)
            row = 8
            indices = [0,1,2,3,4,5,6]
            header = ['Serial #','TCO','Head Count','Gross Amount','Arrears','Union Dues','Net Amount']
            for c in indices:
                sheet.write(row, c, header[c], bold_font)
            
            row += 1
            summary_items = payroll_objs[0].pension_summary_ids
            grand_gross = 0
            staff_strength = 0    
            for summary_item in summary_items:
                sheet.write_number(row, 0, row)
                if summary_item.tco_id:
                    sheet.write_string(row, 1, summary_item.tco_id.name)
                else:
                    sheet.write_string(row, 1, '')
                sheet.write_number(row, 2, summary_item.total_strength, money_format)
                sheet.write_number(row, 3, summary_item.total_gross_income, money_format)
                sheet.write_number(row, 4, summary_item.total_arrears, money_format)
                sheet.write_number(row, 5, summary_item.total_dues, money_format)
                sheet.write_number(row, 6, summary_item.total_net_income, money_format)
                grand_gross += summary_item.total_gross_income
                staff_strength += summary_item.total_strength
                row += 1
                
            #Sum up
            sheet.write_string(row, 0, 'GRAND TOTALS', header_font3)
            for col in [1,2,3,4,5,6]:
                col_name = string.ascii_uppercase[col]
                sheet.write_formula(row, col, '=SUM(' + col_name + '9:' + col_name + str(row) + ')', header_font3_money_format)
            row += 2

            processing_fee = staff_strength * 100
            sheet.write_string(row, 0, 'Gross Pay', header_font3)
            sheet.write_number(row, 1, grand_gross, header_font3_money_format)
            row += 1
            sheet.write_string(row, 0, 'Processing Fees', header_font3)
            sheet.write_number(row, 1, processing_fee, header_font3_money_format)
            row += 1
            sheet.write_string(row, 0, 'Total', header_font3)
            sheet.write_number(row, 1, (processing_fee + grand_gross), header_font3_money_format)
            row += 1
            workbook.close()
            xlsx_data = output.getvalue()
            #payroll_objs[0].update({'pension_exec_summary_report': xlsx_data})
        return xlsx_data    

class payroll_exec_summary_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        header_font = workbook.add_format({'font_size': 18, 'bold': True, 'bottom': 2, 'top': 2})
        header_font2 = workbook.add_format({'font_size': 14, 'bold': True, 'align': 'center', 'bottom': 2, 'top': 2})
        header_font3 = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2})
        header_font3_money_format = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2, 'num_format': '###,###,##0.#0'})
        header_font3_int_format = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2, 'num_format': '###,###,###'})
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        money_format_bold = workbook.add_format({'num_format': '###,###,##0.#0','bold': True})
        format11 = workbook.add_format()
        format11.set_num_format('000')
        
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            _logger.info("Preparing report payroll_exec_summary_report...")
            #Filter item_list based on scenario MDA parameter
            
            sheet = workbook.add_worksheet(payroll_objs[0].name[:31])
            sheet.insert_image('A1:H5', '/odoo/odoo9/osun_ippms.png')
            sheet.merge_range('A6:H7', 'Osun State Executive Staff Summary for: ' + payroll_objs[0].name + ' (' + payroll_objs[0].calendar_id.name + ')', header_font)
            row = 8
            indices = [0,1,2,3,4,5,6,7]
            prev_cal_name = 'N/A'
            if payroll_objs[0].payroll_prev_id:
                prev_cal_name = payroll_objs[0].payroll_prev_id.calendar_id.name
            header = ['SN','Organization',prev_cal_name + ' Strength',payroll_objs[0].calendar_id.name + ' Strength',prev_cal_name + ' Gross',payroll_objs[0].calendar_id.name + ' Gross',prev_cal_name + ' Net',payroll_objs[0].calendar_id.name + ' Net']
            for c in indices:
                sheet.write(row, c, header[c], header_font3)
            row += 1    

            summary_list_current = payroll_objs[0].payroll_summary_ids
            summary_list_previous = False
            if payroll_objs[0].payroll_prev_id:
                summary_list_previous = payroll_objs[0].payroll_prev_id.payroll_summary_ids

            summary_nhf_mda = 0
            summary_paye_mda = 0
            summary_pension_mda = 0
            summary_deduction_other_mda = 0
            summary_gross_mda = 0
            summary_net_mda = 0
            summary_nhf_tescom = 0
            summary_paye_tescom = 0
            summary_pension_tescom = 0
            summary_deduction_other_tescom = 0
            summary_gross_tescom = 0
            summary_net_tescom = 0
            summary_nhf_lth = 0
            summary_paye_lth = 0
            summary_pension_lth = 0
            summary_deduction_other_lth = 0
            summary_gross_lth = 0
            summary_net_lth = 0
            
            totals = [0, 0, 0, 0, 0, 0]
            for summary_item in summary_list_current:
                prev_summary_item = False
                matched_previous_item = False
                if summary_list_previous:
                    matched_previous_item = summary_list_previous.filtered(lambda r: r.department_id == summary_item.department_id)
                if matched_previous_item:
                    prev_summary_item = matched_previous_item
#TODO Uncomment for use in subsequent payroll instances
#                 if summary_item['department_id'].orgtype_id.id in [1,2,3]:
#                     summary_paye_mda += summary_item['total_paye_tax']
#                     summary_nhf_mda += summary_item['total_nhf']
#                     summary_pension_mda += summary_item['total_pension']
#                     summary_deduction_other_mda += summary_item['total_other_deductions']
#                     summary_gross_mda += summary_item['total_gross_income']
#                     summary_net_mda += summary_item['total_net_income']
#                 elif summary_item['department_id'].name.upper().find('TESCOM') > -1:
#                     summary_paye_tescom += summary_item['total_paye_tax']
#                     summary_nhf_tescom += summary_item['total_nhf']
#                     summary_pension_tescom += summary_item['total_pension']
#                     summary_deduction_other_tescom += summary_item['total_other_deductions']
#                     summary_gross_tescom += summary_item['total_gross_income']
#                     summary_net_tescom += summary_item['total_net_income']
#                 elif summary_item['department_id'].name.upper().find('LTH') > -1:
#                     summary_paye_lth += summary_item['total_paye_tax']
#                     summary_nhf_lth += summary_item['total_nhf']
#                     summary_pension_lth += summary_item['total_pension']
#                     summary_deduction_other_lth += summary_item['total_other_deductions']
#                     summary_gross_lth += summary_item['total_gross_income']
#                     summary_net_lth += summary_item['total_net_income']
                
                val1 = 0
                if prev_summary_item and prev_summary_item.department_id == summary_item.department_id:
                    val1 = prev_summary_item.total_strength
                val2 = summary_item.total_strength
                val3 = 0
                if prev_summary_item:
                    val3 = prev_summary_item.total_gross_income
                val4 = summary_item.total_gross_income
                val5 = 0
                if prev_summary_item:
                    val5 = prev_summary_item.total_net_income
                val6 = summary_item.total_net_income
                totals[0] += val1
                totals[1] += val2
                totals[2] += val3
                totals[3] += val4
                totals[4] += val5
                totals[5] += val6
                sheet.write_number(row, 0, (row - 8))
                sheet.write_string(row, 1, summary_item.department_id.name)
                sheet.write_number(row, 2, val1)
                sheet.write_number(row, 3, val2)
                sheet.write_number(row, 4, val3, money_format)
                sheet.write_number(row, 5, val4, money_format)
                sheet.write_number(row, 6, val5, money_format)
                sheet.write_number(row, 7, val6, money_format)
                row += 1

            #Sum up
            sheet.merge_range('A' + str(row) + ':' + 'B' + str(row), 'TOTAL', header_font3)
            for idx in [0,1]:
                sheet.write_number(row - 1, (idx + 2), totals[idx], header_font3_int_format)
            for idx in [2,3,4,5]:
                sheet.write_number(row - 1, (idx + 2), totals[idx], header_font3_money_format)
            row += 1

            #MDA Deductions
            nhf_mda = 0
            paye_mda = 0
            pension_mda = 0
            deduction_other = 0
            gross_mda = 0
            #TESCOM Deductions
            nhf_tescom = 0
            paye_tescom = 0
            pension_tescom = 0
            gross_tescom = 0
            #LTH Deductions
            nhf_lth = 0
            paye_lth = 0
            pension_lth = 0
            gross_lth = 0
            for payroll_item in payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active):
                if payroll_item.active and payroll_item.employee_id.department_id.orgtype_id.id in [1,2,3]:
                    gross_mda += payroll_item.gross_income
                    paye_mda += abs(payroll_item.paye_tax)
                    
                    for item_line in payroll_item.item_line_ids:
                        if 'NHF' in item_line.name:
                            nhf_mda += abs(item_line.amount)
                        elif 'PENSION' in item_line.name:
                            pension_mda += abs(item_line.amount)
                        else:
                            deduction_other += item_line.amount
                elif payroll_item.active and 'TESCOM' in payroll_item.employee_id.department_id.name:
                    gross_tescom += payroll_item.gross_income
                    paye_tescom += abs(payroll_item.paye_tax)
                    
                    for item_line in payroll_item.item_line_ids:
                        if 'NHF' in item_line.name:
                            nhf_tescom += abs(item_line.amount)
                        elif 'PENSION' in item_line.name:
                            pension_tescom += abs(item_line.amount)
                        else:
                            deduction_other += abs(item_line.amount)
                elif payroll_item.active and 'LTH' in payroll_item.employee_id.department_id.name:
                    gross_lth += payroll_item.gross_income
                    paye_lth += abs(payroll_item.paye_tax)
                    
                    for item_line in payroll_item.item_line_ids:
                        if 'NHF' in item_line.name:
                            nhf_lth += abs(item_line.amount)
                        elif 'PENSION' in item_line.name:
                            pension_lth += abs(item_line.amount)
                        else:
                            deduction_other += abs(item_line.amount)

            redemption_bill_mda = gross_mda * 0.05
            redemption_bill_tescom = gross_tescom * 0.05
            redemption_bill_lth = gross_lth * 0.05
            sub_total2_gross = pension_mda + pension_tescom + pension_lth + redemption_bill_mda + redemption_bill_tescom + redemption_bill_lth
            sub_total2_net =  nhf_mda + nhf_tescom + nhf_lth + paye_mda + paye_tescom + paye_lth + pension_mda + pension_tescom + pension_lth + deduction_other

            sheet.merge_range('A' + str(row + 1) + ':' + 'H' + str(row + 1), 'DEDUCTIONS', header_font2)
            row += 1
                                
            sheet.write_string(row, 0, '1')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'NHF', bold_font)
            sheet.write_number(row, 7, (nhf_mda + nhf_tescom + nhf_lth), money_format)
            row += 1
                                
            sheet.write_string(row, 0, '2a')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'PAYE (MDA)', bold_font)
            sheet.write_number(row, 7, paye_mda, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '2b')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'PAYE (TESCOM)', bold_font)
            sheet.write_number(row, 7, paye_tescom, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '2c')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'PAYE (LTH)', bold_font)
            sheet.write_number(row, 7, paye_lth, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '3a')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Contributory Pension (MDA)', bold_font)
            sheet.write_number(row, 7, pension_mda, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '3b')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Contributory Pension (TESCOM)', bold_font)
            sheet.write_number(row, 7, pension_tescom, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '3c')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Contributory Pension (LTH)', bold_font)
            sheet.write_number(row, 7, pension_lth, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '4')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Other Deductions', bold_font)
            sheet.write_number(row, 7, deduction_other, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '5a')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Contributory Pension - Employer (MDA)', bold_font)
            sheet.write_number(row, 6, pension_mda, money_format)
            sheet.write_number(row, 7, pension_mda, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '5b')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Contributory Pension - Employer (TESCOM)', bold_font)
            sheet.write_number(row, 6, pension_tescom, money_format)
            sheet.write_number(row, 7, pension_tescom, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '5c')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Contributory Pension - Employer (LTH)', bold_font)
            sheet.write_number(row, 6, pension_lth, money_format)
            sheet.write_number(row, 7, pension_lth, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '6a')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Redemption Bill - 5% of Wage Bill - Employer (MDA)', bold_font)
            sheet.write_number(row, 6, redemption_bill_mda, money_format)
            sheet.write_number(row, 7, redemption_bill_mda, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '6b')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Redemption Bill - 5% of Wage Bill - Employer (TESCOM)', bold_font)
            sheet.write_number(row, 6, redemption_bill_tescom, money_format)
            sheet.write_number(row, 7, redemption_bill_tescom, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '6c')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Redemption Bill - 5% of Wage Bill - Employer (LTH)', bold_font)
            sheet.write_number(row, 6, redemption_bill_lth, money_format)
            sheet.write_number(row, 7, redemption_bill_lth, money_format)
            row += 1
        
            sheet.write_blank(row, 0, '', header_font3)
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'SUB-TOTAL (DEDUCTIONS)', header_font3)
            sheet.write_number(row, 6, sub_total2_gross, header_font3_money_format)
            sheet.write_number(row, 7, sub_total2_net, header_font3_money_format)
            row += 1
            
            workbook.close()
            xlsx_data = output.getvalue()
            #payroll_objs[0].update({'exec_summary_report': xlsx_data})
#             payroll_objs[0].env.cr.execute("prepare insert_binary_field as insert into ng_state_payroll_payroll (exec_summary_report) values ((decode(encode($1,'HEX'),'HEX')))")
#             payroll_objs[0].env.cr.execute("execute insert_binary_field(%s)", (xlsx_data))
            
            _logger.info("Report payroll_exec_summary_report done.")
        return xlsx_data

class payroll_exec_summary2_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        header_font = workbook.add_format({'font_size': 18, 'bold': True, 'bottom': 2, 'top': 2})
        header_font2 = workbook.add_format({'font_size': 14, 'bold': True, 'align': 'center', 'bottom': 2, 'top': 2})
        bold_font = workbook.add_format({'bold': True})
        header_font3 = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2})
        header_font3_money_format = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2, 'num_format': '###,###,##0.#0'})
        header_font3_int_format = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2, 'num_format': '###,###,###'})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        format11 = workbook.add_format()
        format11.set_num_format('000')
        
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            _logger.info("Preparing report payroll_exec_summary2_report...")
            #Filter item_list based on scenario MDA parameter
            
            sheet = workbook.add_worksheet(payroll_objs[0].name[:31])
            sheet.insert_image('A1:H5', '/odoo/odoo9/osun_ippms.png')
            sheet.merge_range('A6:H7', 'Osun State Executive Staff Summary for: ' + payroll_objs[0].name + ' (' + payroll_objs[0].calendar_id.name + ')', header_font)
            row = 8
            indices = [0,1,2,3,4,5,6,7]
            prev_cal_name = 'N/A'
            if payroll_objs[0].payroll_prev_id:
                prev_cal_name = payroll_objs[0].payroll_prev_id.calendar_id.name
            header = ['SN','Organization',prev_cal_name + ' Strength',payroll_objs[0].calendar_id.name + ' Strength',prev_cal_name + ' Gross',payroll_objs[0].calendar_id.name + ' Gross',prev_cal_name + ' Net',payroll_objs[0].calendar_id.name + ' Net']
            for c in indices:
                sheet.write(row, c, header[c], header_font3)
            row += 1    

            summary_list_current = payroll_objs[0].payroll_summary_ids
            summary_list_previous = False
            if payroll_objs[0].payroll_prev_id:
                summary_list_previous = payroll_objs[0].payroll_prev_id.payroll_summary_ids
                
            summary_nhf_subeb = 0
            summary_paye_subeb = 0
            summary_pension_subeb = 0
            summary_deduction_other_subeb = 0
            summary_gross_subeb = 0
            summary_net_subeb = 0
            summary_nhf_lga = 0
            summary_paye_lga = 0
            summary_pension_lga = 0
            summary_deduction_other_lga = 0
            summary_gross_lga = 0
            summary_net_lga = 0
            
            totals = [0, 0, 0, 0, 0, 0]
            for summary_item in summary_list_current:
                prev_summary_item = False
                matched_previous_item = False
                if summary_list_previous:
                    matched_previous_item = summary_list_previous.filtered(lambda r: r.department_id == summary_item.department_id)
                if matched_previous_item:
                    prev_summary_item = matched_previous_item
#TODO Uncomment for use in subsequent payroll instances
#                 if summary_item['department_id'].orgtype_id.id in [1,2,3]:
#                     summary_paye_mda += summary_item['total_paye_tax']
#                     summary_nhf_mda += summary_item['total_nhf']
#                     summary_pension_mda += summary_item['total_pension']
#                     summary_deduction_other_mda += summary_item['total_other_deductions']
#                     summary_gross_mda += summary_item['total_gross_income']
#                     summary_net_mda += summary_item['total_net_income']
#                 elif summary_item['department_id'].name.upper().find('SUBEB') > -1:
#                     summary_paye_subeb += summary_item['total_paye_tax']
#                     summary_nhf_subeb += summary_item['total_nhf']
#                     summary_pension_subeb += summary_item['total_pension']
#                     summary_deduction_other_subeb += summary_item['total_other_deductions']
#                     summary_gross_subeb += summary_item['total_gross_income']
#                     summary_net_subeb += summary_item['total_net_income']
#                 elif summary_item['department_id'].name.upper().find('LGA') > -1:
#                     summary_paye_lga += summary_item['total_paye_tax']
#                     summary_nhf_lga += summary_item['total_nhf']
#                     summary_pension_lga += summary_item['total_pension']
#                     summary_deduction_other_lga += summary_item['total_other_deductions']
#                     summary_gross_lga += summary_item['total_gross_income']
#                     summary_net_lga += summary_item['total_net_income']
                
                val1 = 0
                if prev_summary_item and prev_summary_item.department_id == summary_item.department_id:
                    val1 = prev_summary_item.total_strength
                val2 = summary_item.total_strength
                val3 = 0
                if prev_summary_item:
                    val3 = prev_summary_item.total_gross_income
                val4 = summary_item.total_gross_income
                val5 = 0
                if prev_summary_item:
                    val5 = prev_summary_item.total_net_income
                val6 = summary_item.total_net_income
                totals[0] += val1
                totals[1] += val2
                totals[2] += val3
                totals[3] += val4
                totals[4] += val5
                totals[5] += val6
                sheet.write_number(row, 0, (row - 8))
                sheet.write_string(row, 1, summary_item.department_id.name)
                sheet.write_number(row, 2, val1)
                sheet.write_number(row, 3, val2)
                sheet.write_number(row, 4, val3, money_format)
                sheet.write_number(row, 5, val4, money_format)
                sheet.write_number(row, 6, val5, money_format)
                sheet.write_number(row, 7, val6, money_format)
                row += 1
            
            row += 1    
            #Sum up
            sheet.merge_range('A' + str(row) + ':' + 'B' + str(row), 'TOTAL', header_font3)
            for idx in [0,1]:
                sheet.write_number(row - 1, (idx + 2), totals[idx], header_font3_int_format)
            for idx in [2,3,4,5]:
                sheet.write_number(row - 1, (idx + 2), totals[idx], header_font3_money_format)
            row += 1

            deduction_other = 0
            #SUBEB Deductions
            nhf_subeb = 0
            paye_subeb = 0
            pension_subeb = 0
            gross_subeb = 0
            #LGA Deductions
            nhf_lga = 0
            paye_lga = 0
            pension_lga = 0
            gross_lga = 0
            for payroll_item in payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active):
                if payroll_item.active and 'SUBEB' in payroll_item.employee_id.department_id.name:
                    gross_subeb += payroll_item.gross_income
                    paye_subeb += abs(payroll_item.paye_tax)
                    
                    for item_line in payroll_item.item_line_ids:
                        if 'NHF' in item_line.name:
                            nhf_subeb += abs(item_line.amount)
                        elif 'PENSION' in item_line.name:
                            pension_subeb += abs(item_line.amount)
                        else:
                            deduction_other += abs(item_line.amount)
                elif payroll_item.active and 'LGA' in payroll_item.employee_id.department_id.name:
                    gross_lga += payroll_item.gross_income
                    paye_lga += abs(payroll_item.paye_tax)
                    
                    for item_line in payroll_item.item_line_ids:
                        if 'NHF' in item_line.name:
                            nhf_lga += abs(item_line.amount)
                        elif 'PENSION' in item_line.name:
                            pension_lga += abs(item_line.amount)
                        else:
                            deduction_other += abs(item_line.amount)

            redemption_bill_subeb = gross_subeb * 0.05
            redemption_bill_lga = gross_lga * 0.05
            sub_total2_gross = pension_subeb + pension_lga + redemption_bill_subeb + redemption_bill_lga
            sub_total2_net =  nhf_subeb + nhf_lga + paye_subeb + paye_lga + pension_subeb + pension_lga + deduction_other

            sheet.merge_range('A' + str(row + 1) + ':' + 'H' + str(row + 1), 'DEDUCTIONS', header_font2)
            row += 1
                                
            sheet.write_string(row, 0, '1')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'NHF', bold_font)
            sheet.write_number(row, 7, (nhf_subeb + nhf_lga), money_format)
            row += 1
                                
            sheet.write_string(row, 0, '2a')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'PAYE (SUBEB)', bold_font)
            sheet.write_number(row, 7, paye_subeb, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '2b')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'PAYE (LGA)', bold_font)
            sheet.write_number(row, 7, paye_lga, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '3a')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Contributory Pension (SUBEB)', bold_font)
            sheet.write_number(row, 7, pension_subeb, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '3b')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Contributory Pension (LGA)', bold_font)
            sheet.write_number(row, 7, pension_lga, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '4')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Other Deductions', bold_font)
            sheet.write_number(row, 7, deduction_other, money_format)
            row += 1
                                                                
            sheet.write_string(row, 0, '5a')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Contributory Pension - Employer (SUBEB)', bold_font)
            sheet.write_number(row, 7, pension_subeb, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '5b')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Contributory Pension - Employer (LGA)', bold_font)
            sheet.write_number(row, 7, pension_lga, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '6a')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Redemption Bill - 5% of Wage Bill - Employer (SUBEB)', bold_font)
            sheet.write_number(row, 7, redemption_bill_subeb, money_format)
            row += 1
                                
            sheet.write_string(row, 0, '6b')
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'Redemption Bill - 5% of Wage Bill - Employer (LGA)', bold_font)
            sheet.write_number(row, 7, redemption_bill_lga, money_format)
            row += 1
        
            sheet.write_blank(row, 0, '', header_font3)
            sheet.merge_range('B' + str(row + 1) + ':' + 'D' + str(row + 1), 'SUB-TOTAL (DEDUCTIONS)', header_font3)
            sheet.write_number(row, 6, sub_total2_gross, header_font3_money_format)
            sheet.write_number(row, 7, sub_total2_net, header_font3_money_format)
            row += 1
            
            workbook.close()
            xlsx_data = output.getvalue()
            #payroll_objs[0].update({'exec_summary2_report': xlsx_data})
            
            _logger.info("Report payroll_exec_summary2_report done.")
        return xlsx_data
    
class payroll_item_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        # Here we will be adding the code to add data
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        header_font3_money_format = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2, 'num_format': '###,###,##0.#0'})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            _logger.info("Creating report...")
            
            item_list = payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active)
    
            sheet = workbook.add_worksheet(payroll_objs[0].name[:31])
            row = 0
            indices = [0,1,2,3,4,5,6,7,8,9,10,11,12]
            header = ['Serial #','Employee Name','Employee Number','Organization','Pay Scheme','Pay Grade','Gross Income','Net Income','PAYE Tax','Pension','Leave Allowance','Unpaid Balance','Bank','Account Number']
            for c in indices:
                sheet.write(row, c, header[c], bold_font)
                
            row = 1
            for payroll_item in item_list:
                sheet.write_number(row, 0, row)
                sheet.write_string(row, 1, payroll_item['employee_id'].name_related)
                sheet.write_string(row, 2, payroll_item['employee_id'].employee_no)
                sheet.write_string(row, 3, payroll_item['employee_id'].department_id.name)
                if payroll_item['employee_id'].payscheme_id:
                    sheet.write_string(row, 4, payroll_item['employee_id'].payscheme_id.name)
                else:
                    sheet.write_string(row, 4, '')
                if payroll_item['employee_id'].level_id:
                    sheet.write_string(row, 5, (str(payroll_item['employee_id'].level_id.paygrade_id.level).zfill(2) + '.' + str(payroll_item['employee_id'].level_id.step).zfill(2)))
                else:
                    sheet.write_string(row, 5, '')
                sheet.write_number(row, 6, payroll_item['gross_income'], money_format)
                sheet.write_number(row, 7, payroll_item['net_income'], money_format)
                sheet.write_number(row, 8, payroll_item['paye_tax'], money_format)
                pension_items = payroll_item['item_line_ids'].filtered(lambda r: r.name.find('PENSION') >= 0)
                if pension_items:
                    pension_total = 0
                    for p in pension_items:
                        pension_total += p.amount
                    sheet.write_number(row, 9, -pension_total, money_format)
                else:
                    sheet.write_number(row, 9, 0, money_format)
                sheet.write_number(row, 10, payroll_item['leave_allowance'], money_format)
                sheet.write_number(row, 11, payroll_item['balance_income'], money_format)
                if payroll_item['employee_id'].bank_id:
                    sheet.write_string(row, 12, payroll_item['employee_id'].bank_id.name)
                else:
                    sheet.write_string(row, 12, '')
                if payroll_item['employee_id'].bank_account_no:
                    sheet.write_string(row, 13, payroll_item['employee_id'].bank_account_no)
                else:
                    sheet.write_string(row, 13, '')                    
                row += 1

            #Sum up
            sheet.write_string(row, 0, 'TOTAL', bold_font)
            for col in [1,2,3,4,5,6,7,8,9,10,11,12]:
                col_name = string.ascii_uppercase[col]
                sheet.write_formula(row, col, '=SUM(' + col_name + '2:' + col_name + str(row) + ')', header_font3_money_format)
                    
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'payroll_report': xlsx_data})
#             param_val = {'bytes': xlsx_data}
#             payroll_objs[0].env.cr.execute("prepare insert_binary_field as insert into ng_state_payroll_payroll (payroll_report) values ((decode(encode($1,'escape'),'escape')))")
#             payroll_objs[0].env.cr.execute("execute insert_binary_field(%s)", (param_val))
            
        return xlsx_data
    
class payroll_paye_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        # Here we will be adding the code to add data
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            _logger.info("Creating report...")
            item_list = payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active)
    
            sheet = workbook.add_worksheet(payroll_objs[0].name[:31])
            row = 0
            indices = [0,1,2,3,4,5,6,7,8,9,10,11]
            header = ['Serial #','Employee Name','Employee Number','Organization','Pay Scheme','Pay Grade','Gross Income','Net Income','Monthly PAYE','Annual PAYE','Bank','Account Number']
            for c in indices:
                sheet.write(row, c, header[c], bold_font)
                
            row = 1    
            for payroll_item in item_list:
                sheet.write_number(row, 0, row)
                sheet.write_string(row, 1, payroll_item['employee_id'].name_related)
                sheet.write_string(row, 2, payroll_item['employee_id'].employee_no)
                sheet.write_string(row, 3, payroll_item['employee_id'].department_id.name)
                if payroll_item['employee_id'].payscheme_id:
                    sheet.write_string(row, 4, payroll_item['employee_id'].payscheme_id.name)
                else:
                    sheet.write_string(row, 4, '')
                if payroll_item['employee_id'].level_id:
                    sheet.write_string(row, 5, (str(payroll_item['employee_id'].level_id.paygrade_id.level).zfill(2) + '.' + str(payroll_item['employee_id'].level_id.step).zfill(2)))
                else:
                    sheet.write_string(row, 5, '')
                sheet.write_number(row, 6, payroll_item['gross_income'], money_format)
                sheet.write_number(row, 7, payroll_item['net_income'], money_format)
                sheet.write_number(row, 8, payroll_item['paye_tax'], money_format)
                sheet.write_number(row, 9, payroll_item['paye_tax_annual'], money_format)
                if payroll_item['employee_id'].bank_id:
                    sheet.write_string(row, 10, payroll_item['employee_id'].bank_id.name)
                else:
                    sheet.write_string(row, 10, '')
                if payroll_item['employee_id'].bank_account_no:
                    sheet.write_string(row, 11, payroll_item['employee_id'].bank_account_no)
                else:
                    sheet.write_string(row, 11, '')
                row += 1

            #Sum up
            sheet.write_string(row, 0, 'TOTAL', bold_font)
            for col in [1,2,3,4,5,6,7,8,9,10,11]:
                col_name = string.ascii_uppercase[col]
                sheet.write_formula(row, col, '=SUM(' + col_name + '2:' + col_name + str(row) + ')', money_format)
                    
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'paye_report': xlsx_data})
        return xlsx_data
            
class pension_item_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        format11 = workbook.add_format()
        format11.set_num_format('000')
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = payroll_objs.pension_item_ids.filtered(lambda r: r.active)

            sheet = workbook.add_worksheet(payroll_objs.name[:31])
            row = 0
            indices = [0,1,2,3,4,5,6,7,8,9,10,11,12,13]
            header = ['Serial #','Retiree Name','Retiree Number','Pension','Gross Income','Arrears','NUP Deductions','HOS Deductions','Net Income','Pension Type','TCO','Bank','Sort Code','Account Number']
            for c in indices:
                sheet.write(row, c, header[c], bold_font)
            row = 1    
            for pension_item in item_list:
                nup = 0
                hos = 0
                for p_item in pension_item.item_line_ids:
                    if 'NUP' in p_item.name:
                        nup = nup + p_item.amount 
                    if 'HOS' in p_item.name:
                        hos = hos + p_item.amount
                sheet.write_number(row, 0, row)
                sheet.write_string(row, 1, pension_item['employee_id'].name_related)
                sheet.write_string(row, 2, pension_item['employee_id'].employee_no)
                sheet.write_number(row, 3, pension_item['employee_id'].annual_pension / 12.0, money_format)
                sheet.write_number(row, 4, pension_item.gross_income, money_format)
                sheet.write_number(row, 5, pension_item.arrears_amount, money_format)
                sheet.write_number(row, 6, nup, money_format)
                sheet.write_number(row, 7, hos, money_format)
                sheet.write_number(row, 8, pension_item.net_income, money_format)
                if pension_item['employee_id'].pensiontype_id:
                    sheet.write_string(row, 9, pension_item['employee_id'].pensiontype_id.name)
                else:
                    sheet.write_string(row, 9, '')
                if pension_item['employee_id'].tco_id:
                    sheet.write_string(row, 10, pension_item['employee_id'].tco_id.name)
                else:
                    sheet.write_string(row, 10, '')
                if pension_item['employee_id'].bank_id:
                    sheet.write_string(row, 11, pension_item['employee_id'].bank_id.name)
                else:
                    sheet.write_string(row, 11, '')
                if pension_item['employee_id'].bank_id:
                    sheet.write_rich_string(row, 12, pension_item['employee_id'].bank_id.bic, format11)
                else:
                    sheet.write_string(row, 12, '')
                if pension_item['employee_id'].bank_account_no:
                    sheet.write_string(row, 13, pension_item['employee_id'].bank_account_no)
                else:
                    sheet.write_string(row, 13, '')
                row += 1

                #Sum up
                sheet.write_string(row, 0, 'TOTAL', bold_font)
                for col in [1,2,3,4,5,6,7,8,9,10,11,12,13]:
                    col_name = string.ascii_uppercase[col]
                    sheet.write_formula(row, col, '=SUM(' + col_name + '2:' + col_name + str(row) + ')', money_format)
                
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'pension_report': xlsx_data})
        return xlsx_data
                 
class pension_tco_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        format11 = workbook.add_format()
        format11.set_num_format('000')
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = payroll_objs[0].pension_item_ids.filtered(lambda r: r.active)

            sheet = None
            row = {}
            header = ['Serial #','Retiree Name','Retiree Number','Pension','Gross Income','Arrears','NUP Deductions','HOS Deductions','Net Income','Pension Type','TCO','Bank','Sort Code','Account Number']
            _logger.info("Header=%s", header) 
            #item_list_filtered = item_list.filtered(lambda r: r.active and r['employee_id'].department_id.orgtype_id.id in [1,2,3]) 
            _logger.info("Item Count=%d", len(item_list)) 
            for pension_item in item_list:
                sheet_name = pension_item['employee_id'].tco_id.name.upper()[:31]
                sheet = workbook.get_worksheet_by_name(sheet_name)
                if sheet is None:
                    sheet = workbook.add_worksheet(sheet_name)
                    row[sheet_name] = 0
                    for i in range(len(header)):
                        sheet.write_string(row[sheet_name], i, header[i], bold_font)
                    row[sheet_name] = 1
                        
                nup = 0
                hos = 0
                arrears = 0
                for p_item in pension_item.item_line_ids:
                    if 'ARREARS' in p_item.name:
                        arrears = arrears + p_item.amount
                    if 'NUP' in p_item.name:
                        nup = nup + p_item.amount 
                    if 'HOS' in p_item.name:
                        hos = hos + p_item.amount
                gross_income = pension_item['gross_income'] + arrears
                net_income = gross_income + nup + hos #deductions are already negative
                sheet.write_number(row[sheet_name], 0, row[sheet_name])
                sheet.write_string(row[sheet_name], 1, pension_item['employee_id'].name_related)
                sheet.write_string(row[sheet_name], 2, pension_item['employee_id'].employee_no)
                sheet.write_number(row[sheet_name], 3, pension_item['gross_income'], money_format)
                sheet.write_number(row[sheet_name], 4, gross_income, money_format)
                sheet.write_number(row[sheet_name], 5, arrears, money_format)
                sheet.write_number(row[sheet_name], 6, nup, money_format)
                sheet.write_number(row[sheet_name], 7, hos, money_format)
                sheet.write_number(row[sheet_name], 8, net_income, money_format)
                if pension_item['employee_id'].pensiontype_id:
                    sheet.write_string(row[sheet_name], 9, pension_item['employee_id'].pensiontype_id.name)
                else:
                    sheet.write_string(row[sheet_name], 9, '')
                if pension_item['employee_id'].tco_id:
                    sheet.write_string(row[sheet_name], 10, pension_item['employee_id'].tco_id.name)
                else:
                    sheet.write_string(row[sheet_name], 10, '')
                if pension_item['employee_id'].bank_id:
                    sheet.write_string(row[sheet_name], 11, pension_item['employee_id'].bank_id.name)
                else:
                    sheet.write_string(row[sheet_name], 11, '')
                if pension_item['employee_id'].bank_id:
                    sheet.write_rich_string(row[sheet_name], 12, pension_item['employee_id'].bank_id.bic, format11)
                else:
                    sheet.write_string(row, 12, '')
                if pension_item['employee_id'].bank_account_no:
                    sheet.write_string(row[sheet_name], 13, pension_item['employee_id'].bank_account_no)
                else:
                    sheet.write_string(row[sheet_name], 13, '')
                row[sheet_name] += 1

            for sheet_name in row:
                sheet = workbook.get_worksheet_by_name(sheet_name)
                #Sum up
                sheet.write_string(row[sheet_name], 0, 'TOTAL', bold_font)
                for col in range(len(header) - 1):
                    col_name = string.ascii_uppercase[(col + 1) % 26]
                    if col > 25:
                        col_name = string.ascii_uppercase[col // 26 - 1] + col_name
                    if col == 25:
                        col_name = 'AA' 
                    sheet.write_formula(row[sheet_name], (col + 1), '=SUM(' + col_name + '2:' + col_name + str(row[sheet_name]) + ')', money_format)
                
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'tco_report': xlsx_data})
        return xlsx_data
            
class payroll_all_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active)

            sheet = None
            row = {}
            for payroll_item in item_list:
                sheet_name = payroll_item['employee_id'].department_id.name[:31]
                sheet = workbook.get_worksheet_by_name(sheet_name)
                if sheet is None:
                    sheet = workbook.add_worksheet(sheet_name)
                    row[sheet_name] = 0
                    indices = [0,1,2,3,4,5,6,7,8,9,10]
                    header = ['Serial #','Employee Name','Employee Number','Organization','Pay Scheme','Pay Grade','Gross Income','Taxable Income','Net Income','PAYE Tax','Unpaid Balance']
                    for c in indices:
                        sheet.write(row[sheet_name], c, header[c], bold_font)
                    row[sheet_name] = 1
                sheet.write_number(row[sheet_name], 0, row[sheet_name])
                sheet.write_string(row[sheet_name], 1, payroll_item['employee_id'].name_related)
                sheet.write_string(row[sheet_name], 2, payroll_item['employee_id'].employee_no)
                sheet.write_string(row[sheet_name], 3, payroll_item['employee_id'].department_id.name)
                if payroll_item['employee_id'].payscheme_id:
                    sheet.write_string(row[sheet_name], 4, payroll_item['employee_id'].payscheme_id.name)
                else:
                    sheet.write_string(row[sheet_name], 4, '')
                if payroll_item['employee_id'].level_id:
                    sheet.write_string(row[sheet_name], 5, (str(payroll_item['employee_id'].level_id.paygrade_id.level).zfill(2) + '.' + str(payroll_item['employee_id'].level_id.step).zfill(2)))
                else:
                    sheet.write_string(row[sheet_name], 5, '')
                sheet.write_number(row[sheet_name], 6, payroll_item['gross_income'], money_format)
                sheet.write_number(row[sheet_name], 7, payroll_item['taxable_income'], money_format)
                sheet.write_number(row[sheet_name], 8, payroll_item['net_income'], money_format)
                sheet.write_number(row[sheet_name], 9, payroll_item['paye_tax'], money_format)
                sheet.write_number(row[sheet_name], 10, payroll_item['balance_income'], money_format)
                row[sheet_name] += 1

            for sheet_name in row:
                sheet = workbook.get_worksheet_by_name(sheet_name)
                #Sum up
                sheet.write_string(row[sheet_name], 0, 'TOTAL', bold_font)
                for col in [1,2,3,4,5,6,7,8,9,10]:
                    col_name = string.ascii_uppercase[col]
                    sheet.write_formula(row[sheet_name], col, '=SUM(' + col_name + '2:' + col_name + str(row[sheet_name]) + ')', money_format)
                
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'departments_report': xlsx_data})
        return xlsx_data
                            
class payroll_tescom_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        money_format_bold = workbook.add_format({'num_format': '###,###,##0.#0','bold': True})
        header_font2 = workbook.add_format({'font_size': 14, 'bold': True})

        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active)

            sheet = None
            row = {}
            header = ['Serial #','Name','Employee Number','Pay Scheme','Pay Grade','Organization','School','Bank','Bank Account']
            payroll_objs[0].env.cr.execute("select distinct name from ng_state_payroll_earning_standard where name is not null and not ltrim(rtrim(name)) = ''")
            headers_select = payroll_objs[0].env.cr.fetchall()
            _logger.info("Fetched Headers=%s", headers_select)
            header_earnings = []
            for item in headers_select:
                header_earnings.append(str(item[0]))
            header.extend(header_earnings)
            header.extend(['Other Earnings','Gross Income','Taxable Income','Net Income','PAYE Tax','Unpaid Balance'])
            _logger.info("Header=%s", header) 
            #item_list_filtered = item_list.filtered(lambda r: r.active and 'TESCOM' in r['employee_id'].department_id.name) 
            _logger.info("Item Count=%d", len(item_list)) 
            # Sort by school and then employee name
            sortable_item_list = []
            school_counter = {}
            for payroll_item in item_list:
                school_name = payroll_item['employee_id'].school_id.name
                if not payroll_item['employee_id'].school_id.name:
                    school_name = ''
                key = payroll_item['employee_id'].department_id.name + '/' + school_name
                if not school_counter.get(key):
                    school_counter[key] = 0
                sortable_item_list.append(payroll_item)
                school_counter[key] += 1
                
            sortable_item_list = sorted(sortable_item_list, key = lambda x: (x['employee_id'].school_id.name, x['employee_id'].name_related))
            
            school_counter_comp = school_counter.copy()
            _logger.info("School Counter=%s", school_counter) 
            
            for payroll_item in sortable_item_list:
                school_name = payroll_item['employee_id'].school_id.name
                if not payroll_item['employee_id'].school_id.name:
                    school_name = ''
                key = payroll_item['employee_id'].department_id.name + '/' + school_name
                # Group by schools
                sheet_name = payroll_item['employee_id'].department_id.name[:31]
                sheet = workbook.get_worksheet_by_name(sheet_name)
                if sheet is None:
                    sheet = workbook.add_worksheet(sheet_name)
                    row[sheet_name] = 0
                    for i in range(len(header)):
                        sheet.write_string(row[sheet_name], i, header[i], bold_font)
                    row[sheet_name] = 1
                        
                col_idx = 0
                # Write School name in row when current school name does not match previous school name
                if payroll_item['employee_id'].school_id and school_counter[key] == school_counter_comp[key]:
                    sheet.merge_range('A' + str(row[sheet_name] + 1) + ':' + 'H' + str(row[sheet_name] + 1), payroll_item['employee_id'].school_id.name, header_font2)
                    row[sheet_name] += 1
                    for i in range(len(header)):
                        sheet.write_string(row[sheet_name], i, header[i], bold_font)
                    row[sheet_name] += 1

                sheet.write_number(row[sheet_name], col_idx, (school_counter[key] - school_counter_comp[key] + 1))
                col_idx += 1
                sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].name_related)
                col_idx += 1
                sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].employee_no)
                col_idx += 1
                if payroll_item['employee_id'].payscheme_id:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].payscheme_id.name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].level_id:
                    sheet.write_string(row[sheet_name], col_idx, (str(payroll_item['employee_id'].level_id.paygrade_id.level).zfill(2) + '.' + str(payroll_item['employee_id'].level_id.step).zfill(2)))
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['department_id']:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['department_id'].name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].school_id:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].school_id.name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].bank_id:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_id.name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].bank_account_no:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_account_no)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                for h in header_earnings:
                    line_item = payroll_item['item_line_ids'].filtered(lambda r: r.name == h)
                    if line_item:
                        sheet.write_number(row[sheet_name], col_idx, line_item[0].amount, money_format)
                    else:
                        sheet.write_number(row[sheet_name], col_idx, 0, money_format)
                    col_idx += 1

                #TODO Sum all nonstandard earnings - use prefix 'OTHER EARNINGS - '
                other_earnings =  payroll_item['item_line_ids'].filtered(lambda r: r.name.startswith('OTHER EARNINGS - '))
                other_earnings_total = 0
                for o in other_earnings:
                    other_earnings_total += o.amount          
                sheet.write_number(row[sheet_name], col_idx, other_earnings_total, money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['gross_income'], money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['taxable_income'], money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['net_income'], money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['paye_tax'], money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['balance_income'], money_format)
                row[sheet_name] += 1
                
                if school_counter_comp[key] > 0:
                    school_counter_comp[key] -= 1
                    
                if school_counter_comp[key] == 0:
                    #TODO write sub-totaling column for the school group
                    sheet.write_string(row[sheet_name], 0, 'TOTAL', bold_font)
                    for col in range(len(header) - 1):
                        col_name = string.ascii_uppercase[(col + 1) % 26]
                        if col > 25:
                            col_name = string.ascii_uppercase[col // 26 - 1] + col_name
                        if col == 25:
                            col_name = 'AA' 
                        if col == 51:
                            col_name = 'BA' 
                        sheet.write_formula(row[sheet_name], (col + 1), '=SUM(' + col_name + str(row[sheet_name] - school_counter[key] + 1) + ':' + col_name + str(row[sheet_name]) + ')', money_format_bold)
                    row[sheet_name] += 1
                    row[sheet_name] += 1
                    
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'tescom_report': xlsx_data})
        return xlsx_data
                            
class payroll_tescom_school_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            _logger.info("Creating report...payroll_tescom_school_report")

            payroll_objs[0].env.cr.execute("select distinct department_id from ng_state_payroll_payroll_item where active='t' and payroll_id=" + str(payroll_objs[0].id))
            department_ids = payroll_objs[0].env.cr.fetchall()

            sheet = workbook.add_worksheet(payroll_objs[0].name[:31])
            row = 0
            indices = [0,1,2,3,4,5,6]
            header = ['Serial #','School','Strength','PAYE','Pension','Gross Income','Net Pay']
            for c in indices:
                sheet.write(row, c, header[c], bold_font)
            
            row = 1
            for dept_id in department_ids:
                summary_items = payroll_objs[0].payroll_schoolsummary_ids.filtered(lambda r: r.school_id.org_id.id == dept_id[0])
                payroll_objs[0].env.cr.execute("select name from hr_department where id=" + str(dept_id[0]))
                department_names = payroll_objs[0].env.cr.fetchone()
                sheet.merge_range('A' + str(row + 1) + ':' + 'G' + str(row + 1), department_names[0], bold_font)
                row += 1
                for summary_item in summary_items:
                    sheet.write_number(row, 0, dept_id[0])
                    if summary_item.school_id:
                        sheet.write_string(row, 1, summary_item.school_id.name)
                    else:
                        sheet.write_string(row, 1, '')
                    sheet.write_number(row, 2, summary_item.total_strength)
                    sheet.write_number(row, 3, summary_item.total_paye_tax, money_format)
                    sheet.write_number(row, 4, summary_item.total_pension, money_format)
                    sheet.write_number(row, 5, summary_item.total_gross_income, money_format)
                    sheet.write_number(row, 6, summary_item.total_net_income, money_format)
                    row += 1
                    
                #Sum up
                sheet.write_string(row, 1, 'TOTAL', bold_font)
                for col in [2,3,4,5,6]:
                    col_name = string.ascii_uppercase[col]
                    if col > 2:
                        sheet.write_formula(row, col, '=SUM(' + col_name + str(row + 1 - len(summary_items)) + ':' + col_name + str(row) + ')', money_format)
                    else:
                        sheet.write_formula(row, col, '=SUM(' + col_name + str(row + 1 - len(summary_items)) + ':' + col_name + str(row) + ')')
                        
                row += 2

            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'tescom_school_report': xlsx_data})
        return xlsx_data

            
class payroll_leavebonus_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        header_font = workbook.add_format({'font_size': 18, 'bold': True, 'bottom': 2, 'top': 2})
        header_font2 = workbook.add_format({'font_size': 14, 'bold': True, 'align': 'center', 'bottom': 2, 'top': 2})
        header_font3 = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2})
        header_font3_money_format = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2, 'num_format': '###,###,##0.#0'})
        header_font3_int_format = workbook.add_format({'bold': True, 'bottom': 2, 'top': 2, 'num_format': '###,###,###'})
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        money_format_bold = workbook.add_format({'num_format': '###,###,##0.#0','bold': True})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active and r.leave_allowance > 0.0)
            summary_list = payroll_objs[0].payroll_summary_ids
            
            #Add front-cover worksheet
            cover_sheet = workbook.add_worksheet('SUMMARY')
            
            cover_sheet.insert_image('A1:H5', '/odoo/odoo9/osun_ippms.png')
            cover_sheet.merge_range('A6:H7', 'Leave Allowance: ' + payroll_objs[0].name + ' (' + payroll_objs[0].calendar_id.name + ')', header_font)
            cover_row = 8
            indices = [0,1,2]
            header = ['Organization',payroll_objs[0].calendar_id.name + ' Strength','Leave Allowance']
            for c in indices:
                cover_sheet.write(cover_row, c, header[c], header_font3)
            cover_row += 1    
            for summary_item in summary_list:
                cover_sheet.write_string(cover_row, 0, summary_item.department_id.name)
                cover_sheet.write_number(cover_row, 1, len(item_list.filtered(lambda r: r.employee_id.department_id.id == summary_item.department_id.id)))
                cover_sheet.write_number(cover_row, 2, summary_item.total_leave_allowance, money_format)
                cover_row += 1
            cover_sheet.write_string(cover_row, 0, 'GRAND TOTAL', header_font3)
            for col in [1,2]:
                col_name = string.ascii_uppercase[col]
                if col == 1:
                    cover_sheet.write_formula(cover_row, col, '=SUM(' + col_name + '9:' + col_name + str(cover_row) + ')', money_format)
                else:
                    cover_sheet.write_formula(cover_row, col, '=SUM(' + col_name + '9:' + col_name + str(cover_row) + ')', money_format)
                    
            sheet = None
            row = {}
            for payroll_item in item_list:
                sheet_name = payroll_item['employee_id'].department_id.name[:31]
                sheet = workbook.get_worksheet_by_name(sheet_name)
                if sheet is None:
                    sheet = workbook.add_worksheet(sheet_name)
                    row[sheet_name] = 0
                    indices = [0,1,2,3,4,5,6,7,8,9,10]
                    header = ['Serial #','Employee Name','Employee Number','Organization','Pay Scheme','Pay Grade','Basic','Leave Allowance','Bank','Sort Code','Account Number']
                    for c in indices:
                        sheet.write(row[sheet_name], c, header[c], bold_font)
                    row[sheet_name] = 1
                #If at least 6 months worked
                emp = payroll_item['employee_id']

                if payroll_item['leave_allowance'] > 0.0:                    
                    sheet.write_number(row[sheet_name], 0, row[sheet_name])
                    sheet.write_string(row[sheet_name], 1, emp.name_related)
                    sheet.write_string(row[sheet_name], 2, emp.employee_no)
                    sheet.write_string(row[sheet_name], 3, emp.department_id.name)
                    if emp.payscheme_id:
                        sheet.write_string(row[sheet_name], 4, emp.payscheme_id.name)
                    else:
                        sheet.write_string(row[sheet_name], 4, '')
                    if emp.level_id:
                        sheet.write_string(row[sheet_name], 5, (str(emp.level_id.paygrade_id.level).zfill(2) + '.' + str(emp.level_id.step).zfill(2)))
                    else:
                        sheet.write_string(row[sheet_name], 5, '')
                    basic_salary = emp.standard_earnings.filtered(lambda r: r.active == True and r.name == 'BASIC SALARY')
                    if basic_salary:
                        sheet.write_number(row[sheet_name], 6, (basic_salary.amount / 12), money_format)
                    else:
                        sheet.write_number(row[sheet_name], 6, 0.0, money_format)
                    sheet.write_number(row[sheet_name], 7, payroll_item['leave_allowance'], money_format)
                    if emp.bank_id:
                        sheet.write_string(row[sheet_name], 8, emp.bank_id.name)
                    else:
                        sheet.write_string(row[sheet_name], 8, '')
                    if emp.bank_id.bic:
                        sheet.write_string(row[sheet_name], 9, emp.bank_id.bic[:3])
                    else:
                        sheet.write_string(row[sheet_name], 9, '')
                    if emp.bank_account_no:
                        sheet.write_string(row[sheet_name], 10, emp.bank_account_no)
                    else:
                        sheet.write_string(row[sheet_name], 10, '')
                    
                    row[sheet_name] += 1

            for sheet_name in row:
                sheet = workbook.get_worksheet_by_name(sheet_name)
                #Sum up
                sheet.write_string(row[sheet_name], 0, 'TOTAL', bold_font)
                for col in [1,2,3,4,5,6,7,8,9,10]:
                    col_name = string.ascii_uppercase[col]
                    sheet.write_formula(row[sheet_name], col, '=SUM(' + col_name + '2:' + col_name + str(row[sheet_name]) + ')', money_format)
                
            workbook.close()

            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'leavebonus_report': xlsx_data})
        return xlsx_data
                
class payroll_mda_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active)

            sheet = None
            row = {}
            header = ['Serial #','Name','Employee Number','Pay Scheme','Pay Grade','MDA','Bank','Bank Account']
            payroll_objs[0].env.cr.execute("select distinct name from ng_state_payroll_earning_standard where name is not null and not ltrim(rtrim(name)) = ''")
            headers_select = payroll_objs[0].env.cr.fetchall()
            _logger.info("Fetched Headers=%s", headers_select)
            header_earnings = []
            for item in headers_select:
                header_earnings.append(str(item[0]))
            header.extend(header_earnings)
            header.extend(['Other Earnings','Gross Income','Taxable Income','Net Income','PAYE Tax','Unpaid Balance'])
            _logger.info("Header=%s", header) 
            #item_list_filtered = item_list.filtered(lambda r: r.active and r['employee_id'].department_id.orgtype_id.id in [1,2,3]) 
            _logger.info("Item Count=%d", len(item_list)) 
            for payroll_item in item_list:
                sheet_name = payroll_item['employee_id'].department_id.name[:31]
                sheet = workbook.get_worksheet_by_name(sheet_name)
                if sheet is None:
                    sheet = workbook.add_worksheet(sheet_name)
                    row[sheet_name] = 0
                    for i in range(len(header)):
                        sheet.write_string(row[sheet_name], i, header[i], bold_font)
                    row[sheet_name] = 1
                        
                col_idx = 0
                sheet.write_number(row[sheet_name], col_idx, row[sheet_name])
                col_idx += 1
                sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].name_related)
                col_idx += 1
                sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].employee_no)
                col_idx += 1
                if payroll_item['employee_id'].payscheme_id:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].payscheme_id.name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].level_id:
                    sheet.write_string(row[sheet_name], col_idx, (str(payroll_item['employee_id'].level_id.paygrade_id.level).zfill(2) + '.' + str(payroll_item['employee_id'].level_id.step).zfill(2)))
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['department_id']:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['department_id'].name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].bank_id:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_id.name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].bank_account_no:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_account_no)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                for h in header_earnings:
                    line_item = payroll_item['item_line_ids'].filtered(lambda r: r.name == h)
                    if line_item:
                        sheet.write_number(row[sheet_name], col_idx, line_item[0].amount, money_format)
                    else:
                        sheet.write_number(row[sheet_name], col_idx, 0, money_format)
                    col_idx += 1

                #TODO Sum all nonstandard earnings - use prefix 'OTHER EARNINGS - '
                other_earnings =  payroll_item['item_line_ids'].filtered(lambda r: r.name.startswith('OTHER EARNINGS - '))
                other_earnings_total = 0
                for o in other_earnings:
                    other_earnings_total += o.amount          
                sheet.write_number(row[sheet_name], col_idx, other_earnings_total, money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['gross_income'], money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['taxable_income'], money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['net_income'], money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['paye_tax'], money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['balance_income'], money_format)
                row[sheet_name] += 1

                for sheet_name in row:
                    sheet = workbook.get_worksheet_by_name(sheet_name)
                    #Sum up
                    sheet.write_string(row[sheet_name], 0, 'TOTAL', bold_font)
                    for col in range(len(header) - 1):
                        col_name = string.ascii_uppercase[(col + 1) % 26]
                        if col > 25:
                            col_name = string.ascii_uppercase[col // 26 - 1] + col_name
                        if col == 25:
                            col_name = 'AA' 
                        sheet.write_formula(row[sheet_name], (col + 1), '=SUM(' + col_name + '2:' + col_name + str(row[sheet_name]) + ')', money_format)
                
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'mda_report': xlsx_data})
        return xlsx_data
            
class payroll_mda_summarized_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active)

            sheet = None
            row = {}
            header = ['Serial #','Name','Employee Number','Pay Scheme','Pay Grade','MDA','Bank','Bank Account']
            header_tail = ['Other Earnings','Gross','Net','PAYE']
            payroll_objs[0].env.cr.execute("select distinct name from ng_state_payroll_earning_standard where active='t' and name is not null and not ltrim(rtrim(name)) = ''")
            headers_select = payroll_objs[0].env.cr.fetchall()
            header_earnings = []
            _logger.info("Fetched Headers=%s", headers_select)
            for item in headers_select:
                header_earnings.append(str(item[0]))
                
            header.extend(header_earnings)
            header.extend(header_tail)
            _logger.info("Header=%s", header) 
            #item_list_filtered = item_list.filtered(lambda r: r.active and r['employee_id'].department_id.orgtype_id.id in [1,2,3]) 
            _logger.info("Item Count=%d", len(item_list)) 
            for payroll_item in item_list:
                sheet_name = payroll_item['employee_id'].department_id.name[:31]
                sheet = workbook.get_worksheet_by_name(sheet_name)
                if sheet is None:
                    sheet = workbook.add_worksheet(sheet_name)
                    row[sheet_name] = 0
                    for i in range(len(header)):
                        sheet.write_string(row[sheet_name], i, header[i], bold_font)
                    row[sheet_name] = 1
                        
                col_idx = 0
                sheet.write_number(row[sheet_name], col_idx, row[sheet_name])
                col_idx += 1
                sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].name_related)
                col_idx += 1
                sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].employee_no)
                col_idx += 1
                if payroll_item['employee_id'].payscheme_id:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].payscheme_id.name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].level_id:
                    sheet.write_string(row[sheet_name], col_idx, (str(payroll_item['employee_id'].level_id.paygrade_id.level).zfill(2) + '.' + str(payroll_item['employee_id'].level_id.step).zfill(2)))
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['department_id']:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['department_id'].name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].bank_id:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_id.name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].bank_account_no:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_account_no)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                for h in header_earnings:
                    line_item = payroll_item['item_line_ids'].filtered(lambda r: r.name == h)
                    if line_item:
                        sheet.write_number(row[sheet_name], col_idx, line_item[0].amount, money_format)
                    else:
                        sheet.write_number(row[sheet_name], col_idx, 0, money_format)
                    col_idx += 1

                #TODO Sum all nonstandard earnings - use prefix 'OTHER EARNINGS - '
                other_earnings =  payroll_item['item_line_ids'].filtered(lambda r: r.name.startswith('OTHER EARNINGS - '))
                other_earnings_total = 0
                for o in other_earnings:
                    other_earnings_total += o.amount          
                sheet.write_number(row[sheet_name], col_idx, other_earnings_total, money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['gross_income'], money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['net_income'], money_format)
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, payroll_item['paye_tax'], money_format)
                row[sheet_name] += 1

            for sheet_name in row:
                sheet = workbook.get_worksheet_by_name(sheet_name)
                #Sum up
                sheet.write_string(row[sheet_name], 0, 'TOTAL', bold_font)
                for col in range(len(header) - 1):
                    col_name = string.ascii_uppercase[(col + 1) % 26]
                    if col > 25:
                        col_name = string.ascii_uppercase[col // 26 - 1] + col_name
                    if col == 25:
                        col_name = 'AA' 
                    sheet.write_formula(row[sheet_name], (col + 1), '=SUM(' + col_name + '2:' + col_name + str(row[sheet_name]) + ')', money_format)
                
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'mda_summary_report': xlsx_data})
        return xlsx_data
            
class payroll_mda_deduction_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active)

            sheet = None
            row = {}
            header = ['Serial #','Name','Employee Number','Pay Scheme','Pay Grade','Organization','Bank','Bank Account','Gross Income','Net Income']
            payroll_objs[0].env.cr.execute("select distinct id from ng_state_payroll_payroll_item where active='t' and payroll_id=" + str(payroll_objs[0].id))
            item_ids_fetched = payroll_objs[0].env.cr.fetchall()
            item_ids = []
            for e in item_ids_fetched:
                item_ids.append(str(e[0]))
            
            header_deductions = []
            if len(item_ids) > 0:
                payroll_objs[0].env.cr.execute("select distinct name from ng_state_payroll_payroll_item_line where amount < 0 and name not like '%PENSION%' and item_id in (" + ",".join(item_ids) + ") and name is not null and not ltrim(rtrim(name)) = ''")
                headers_select = payroll_objs[0].env.cr.fetchall()
                _logger.info("Fetched Headers=%s", headers_select)
                for item in headers_select:
                    header_deductions.append(str(item[0].replace('OTHER DEDUCTIONS - ', '')))
                
                header_deductions.append('Pension')    
                header.extend(header_deductions)
                _logger.info("Header=%s", header) 
                for payroll_item in item_list:
                    sheet_name = payroll_item['employee_id'].department_id.name[:31]
                    sheet = workbook.get_worksheet_by_name(sheet_name)
                    if sheet is None:
                        sheet = workbook.add_worksheet(sheet_name)
                        row[sheet_name] = 0
                        for i in range(len(header)):
                            sheet.write_string(row[sheet_name], i, header[i], bold_font)
                        row[sheet_name] = 1
                            
                    col_idx = 0
                    sheet.write_number(row[sheet_name], col_idx, row[sheet_name])
                    col_idx += 1
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].name_related)
                    col_idx += 1
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].employee_no)
                    col_idx += 1
                    if payroll_item['employee_id'].payscheme_id:
                        sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].payscheme_id.name)
                    else:
                        sheet.write_string(row[sheet_name], col_idx, '')
                    col_idx += 1
                    if payroll_item['employee_id'].level_id:
                        sheet.write_string(row[sheet_name], col_idx, (str(payroll_item['employee_id'].level_id.paygrade_id.level).zfill(2) + '.' + str(payroll_item['employee_id'].level_id.step).zfill(2)))
                    else:
                        sheet.write_string(row[sheet_name], col_idx, '')
                    col_idx += 1
                    if payroll_item['department_id']:
                        sheet.write_string(row[sheet_name], col_idx, payroll_item['department_id'].name)
                    else:
                        sheet.write_string(row[sheet_name], col_idx, '')
                    col_idx += 1
                    if payroll_item['employee_id'].bank_id:
                        sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_id.name)
                    else:
                        sheet.write_string(row[sheet_name], col_idx, '')
                    col_idx += 1
                    if payroll_item['employee_id'].bank_account_no:
                        sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_account_no)
                    else:
                        sheet.write_string(row[sheet_name], col_idx, '')
                    col_idx += 1
                    sheet.write_number(row[sheet_name], col_idx, payroll_item['gross_income'], money_format)
                    col_idx += 1
                    sheet.write_number(row[sheet_name], col_idx, payroll_item['net_income'], money_format)
                    col_idx += 1
                    for h in header_deductions:
                        if h == 'Pension':
                            pension_items = payroll_item['item_line_ids'].filtered(lambda r: r.name.find('PENSION') >= 0)
                            if pension_items:
                                pension_total = 0
                                for p in pension_items:
                                    pension_total += p.amount
                                sheet.write_number(row[sheet_name], col_idx, -pension_total, money_format)
                            else:
                                sheet.write_number(row[sheet_name], col_idx, 0, money_format)
                        else:
                            line_item = payroll_item['item_line_ids'].filtered(lambda r: r.name.replace('OTHER DEDUCTIONS - ', '') == h)
                            if line_item:
                                sheet.write_number(row[sheet_name], col_idx, -line_item[0].amount, money_format)
                            else:
                                sheet.write_number(row[sheet_name], col_idx, 0, money_format)
                            
                        col_idx += 1
        
                    row[sheet_name] += 1
    
                for sheet_name in row:
                    sheet = workbook.get_worksheet_by_name(sheet_name)
                    #Sum up
                    sheet.write_string(row[sheet_name], 0, 'TOTAL', bold_font)
                    for col in range(len(header) - 1):
                        col_name = string.ascii_uppercase[(col + 1) % 26]
                        if col > 25:
                            col_name = string.ascii_uppercase[col // 26 - 1] + col_name
                        if col == 25:
                            col_name = 'AA' 
                        sheet.write_formula(row[sheet_name], (col + 1), '=SUM(' + col_name + '2:' + col_name + str(row[sheet_name]) + ')', money_format)
                    
                workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'mda_deduction_report': xlsx_data})
        return xlsx_data
            
class payroll_mda_deduction_head_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active)

            sheet = None
            row = {}
            for payroll_item in item_list:
                for deduction in payroll_item.item_line_ids.filtered(lambda r: r.amount < 0 and not r.name.find('PENSION') >= 0):
                    sheet_name = deduction.name.replace('OTHER DEDUCTIONS - ', '')[:31]
                    header = ['Serial #','Name','Employee Number','Pay Scheme','Pay Grade','MDA','Bank','Bank Account',sheet_name]
                    sheet = workbook.get_worksheet_by_name(sheet_name)
                    if sheet is None:
                        sheet = workbook.add_worksheet(sheet_name)
                        row[sheet_name] = 0
                        for i in range(len(header)):
                            sheet.write_string(row[sheet_name], i, header[i], bold_font)
                        row[sheet_name] = 1
                    col_idx = 0
                    sheet.write_number(row[sheet_name], col_idx, row[sheet_name])
                    col_idx += 1
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].name_related)
                    col_idx += 1
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].employee_no)
                    col_idx += 1
                    if payroll_item['employee_id'].payscheme_id:
                        sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].payscheme_id.name)
                    else:
                        sheet.write_string(row[sheet_name], col_idx, '')
                    col_idx += 1
                    if payroll_item['employee_id'].level_id:
                        sheet.write_string(row[sheet_name], col_idx, (str(payroll_item['employee_id'].level_id.paygrade_id.level).zfill(2) + '.' + str(payroll_item['employee_id'].level_id.step).zfill(2)))
                    else:
                        sheet.write_string(row[sheet_name], col_idx, '')
                    col_idx += 1
                    if payroll_item['department_id']:
                        sheet.write_string(row[sheet_name], col_idx, payroll_item['department_id'].name)
                    else:
                        sheet.write_string(row[sheet_name], col_idx, '')
                    col_idx += 1
                    if payroll_item['employee_id'].bank_id:
                        sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_id.name)
                    else:
                        sheet.write_string(row[sheet_name], col_idx, '')
                    col_idx += 1
                    if payroll_item['employee_id'].bank_account_no:
                        sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_account_no)
                    else:
                        sheet.write_string(row[sheet_name], col_idx, '')
                    col_idx += 1
                    sheet.write_number(row[sheet_name], col_idx, -deduction.amount, money_format)
                
                    row[sheet_name] += 1
                
                pension = 0
                for deduction in payroll_item.item_line_ids.filtered(lambda r: r.amount < 0 and r.name.find('PENSION') >= 0):
                    pension += deduction.amount
                sheet_name = 'Pension'
                header = ['Serial #','Name','Employee Number','Pay Scheme','Pay Grade','MDA','Bank','Bank Account',sheet_name]
                sheet = workbook.get_worksheet_by_name(sheet_name)
                if sheet is None:
                    sheet = workbook.add_worksheet(sheet_name)
                    row[sheet_name] = 0
                    for i in range(len(header)):
                        sheet.write_string(row[sheet_name], i, header[i], bold_font)
                    row[sheet_name] = 1
                col_idx = 0
                sheet.write_number(row[sheet_name], col_idx, row[sheet_name])
                col_idx += 1
                sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].name_related)
                col_idx += 1
                sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].employee_no)
                col_idx += 1
                if payroll_item['employee_id'].payscheme_id:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].payscheme_id.name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].level_id:
                    sheet.write_string(row[sheet_name], col_idx, (str(payroll_item['employee_id'].level_id.paygrade_id.level).zfill(2) + '.' + str(payroll_item['employee_id'].level_id.step).zfill(2)))
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['department_id']:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['department_id'].name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].bank_id:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_id.name)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                if payroll_item['employee_id'].bank_account_no:
                    sheet.write_string(row[sheet_name], col_idx, payroll_item['employee_id'].bank_account_no)
                else:
                    sheet.write_string(row[sheet_name], col_idx, '')
                col_idx += 1
                sheet.write_number(row[sheet_name], col_idx, -pension, money_format)
            
                row[sheet_name] += 1
                
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'mda_deduction_head_report': xlsx_data})
        return xlsx_data
                            
class pension_mda_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, payroll_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = payroll_objs[0].payroll_item_ids.filtered(lambda r: r.active)

            sheet = None
            row = {}    
            for payroll_item in item_list:
                sheet_name = payroll_item['employee_id'].department_id.name[:31]
                sheet = workbook.get_worksheet_by_name(sheet_name)
                if sheet is None:
                    sheet = workbook.add_worksheet(sheet_name)
                    row[sheet_name] = 0
                    indices = [0,1,2,3,4,5,6,7,8]
                    header = ['Serial #','Employee Name','Employee Number','Pay Scheme','Pay Grade','MDA','Pension Total','Pension PIN','PFA Name']
                    for c in indices:
                        sheet.write(row[sheet_name], c, header[c], bold_font)
                    row[sheet_name] = 1

                #TODO Sum all standard pension deductions
                pension_total = 0
                for deduction in payroll_item.item_line_ids.filtered(lambda r: r.amount < 0 and r.name.find('PENSION') >= 0):
                    pension_total += deduction.amount

                sheet.write_number(row[sheet_name], 0, row[sheet_name])
                sheet.write_string(row[sheet_name], 1, payroll_item['employee_id'].name_related)
                sheet.write_string(row[sheet_name], 2, payroll_item['employee_id'].employee_no)
                sheet.write_string(row[sheet_name], 3, payroll_item['employee_id'].payscheme_id.name)
                if payroll_item['employee_id'].level_id:
                    sheet.write_string(row[sheet_name], 4, (str(payroll_item['employee_id'].level_id.paygrade_id.level).zfill(2) + '.' + str(payroll_item['employee_id'].level_id.step).zfill(2)))
                else:
                    sheet.write_string(row[sheet_name], 4, '')
                sheet.write_string(row[sheet_name], 5, payroll_item['employee_id'].department_id.name)
                sheet.write_number(row[sheet_name], 6, pension_total, money_format)
                if payroll_item['employee_id'].sinid:
                    sheet.write_string(row[sheet_name], 7, payroll_item['employee_id'].sinid)
                else:
                    sheet.write_string(row[sheet_name], 7, '')
                if payroll_item['employee_id'].pfa_id:
                    sheet.write_string(row[sheet_name], 8, payroll_item['employee_id'].pfa_id.name)
                else:
                    sheet.write_string(row[sheet_name], 8, '')
                row[sheet_name] += 1

            for sheet_name in row:
                sheet = workbook.get_worksheet_by_name(sheet_name)
                #Sum up
                sheet.write_string(row[sheet_name], 0, 'TOTAL', bold_font)
                for col in [1,2,3,4,5,6,7,8]:
                    col_name = string.ascii_uppercase[col]
                    sheet.write_formula(row[sheet_name], col, '=SUM(' + col_name + '2:' + col_name + str(row[sheet_name]) + ')', money_format)
                
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #payroll_objs[0].update({'pension_mda_report': xlsx_data})
        return xlsx_data

class payment_item_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, scenario_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = []
            if scenario_objs[0].scenario_item_ids:
                item_list = scenario_objs[0].payment_ids.filtered(lambda r: r.active)
            elif scenario_objs[0].scenario2_item_ids:
                item_list = scenario_objs[0].payment2_ids.filtered(lambda r: r.active)

            sheet = workbook.add_worksheet(scenario_objs[0].name[:31])
            row = 0
            indices = [0,1,2,3,4,5,6,7,8,9]
            header = ['Serial #', 'Employee Name','Employee Number','MDA/TCO','Bank','Bank Account','Net Income (100%)','Payment Amount','Payment Balance','Percentage']
            for c in indices:
                sheet.write(row, c, header[c], bold_font)
            
            row = 1    
            for payment_item in item_list:
                sheet.write_number(row, 0, row)
                sheet.write_string(row, 1, payment_item['employee_id'].name_related)
                sheet.write_string(row, 2, payment_item['employee_id'].employee_no)
                if payment_item['employee_id'].department_id:
                    sheet.write_string(row, 3, payment_item['employee_id'].department_id.name)
                elif payment_item['employee_id'].tco_id:
                    sheet.write_string(row, 3, payment_item['employee_id'].tco_id.name)
                else:
                    sheet.write_string(row, 3, '')
                if payment_item['employee_id'].bank_id:
                    sheet.write_string(row, 4, payment_item['employee_id'].bank_id.name)
                else:
                    sheet.write_string(row, 4, '')
                if payment_item['employee_id'].bank_account_no:
                    sheet.write_string(row, 5, payment_item['employee_id'].bank_account_no)
                else:
                    sheet.write_string(row, 5, '')
                sheet.write_number(row, 6, payment_item['net_income'], money_format)
                sheet.write_number(row, 7, payment_item['amount'], money_format)
                sheet.write_number(row, 8, payment_item['balance_income'], money_format)
                sheet.write_number(row, 9, payment_item['percentage'], money_format)
                row += 1

            #Sum up
            sheet.write_string(row, 0, 'TOTAL', bold_font)
            for col in [1,2,3,4,5,6,7,8,9]:
                col_name = string.ascii_uppercase[col]
                    
                sheet.write_formula(row, col, '=SUM(' + col_name + '2:' + col_name + str(row) + ')', money_format)
               
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #scenario_objs[0].update({'employee_report': xlsx_data})
        return xlsx_data

class payment_nibbs_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, scenario_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format_string = {'num_format': '########0.#0'}
        money_format = workbook.add_format(money_format_string)
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            item_list = []
            if scenario_objs[0].payroll_id.do_payroll:
                item_list = scenario_objs[0].payment_ids.filtered(lambda r: r.active)
            else:
                item_list = scenario_objs[0].payment2_ids.filtered(lambda r: r.active)

            sheet = workbook.add_worksheet(scenario_objs[0].name[:31])
            row = 0
            indices = [0,1,2,3,4,5]
            header = ['Serial Number','Account Number','Bank Code','Amount','Account Name','Narration']
            for c in indices:
                sheet.write(row, c, header[c], bold_font)
            
            row = 1    
            for payment_item in item_list:
                sheet.write_number(row, 0, row)
                sheet.write_string(row, 1, payment_item['employee_id'].bank_account_no)
                sheet.write_string(row, 2, payment_item['employee_id'].bank_id.bic[:3])
                sheet.write_number(row, 3, payment_item['amount'], money_format)
                sheet.write_string(row, 4, payment_item['employee_id'].name_related)
                sheet.write_string(row, 5, str(int(payment_item['percentage'])) + "p for " + payment_item['scenario_id'].payroll_id.name + ", " + payment_item['scenario_id'].payroll_id.calendar_id.name + " " + payment_item['employee_id'].employee_no)
                row += 1
                
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #scenario_objs[0].update({'nibbs_report': xlsx_data})
        return xlsx_data

class deduction_nibbs_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, scenario_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format_string = {'num_format': '########0.#0'}
        money_format = workbook.add_format(money_format_string)
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            sheet = workbook.add_worksheet(scenario_objs[0].name[:31])
            row = 0
            indices = [0,1,2,3,4,5]
            header = ['Serial Number','Account Number','Bank Code','Amount','Beneficiary Name','Narration']
            for c in indices:
                sheet.write(row, c, header[c], bold_font)
            
            row += 1
            sheet.write_number(row, 0, row)
            sheet.write_string(row, 1, '8000000001')
            sheet.write_string(row, 2, '011')
            sheet.write_number(row, 3, (scenario_objs[0].payroll_id.total_gross_payroll - scenario_objs[0].payroll_id.total_net_pension - scenario_objs[0].payroll_id.total_tax_payroll), money_format)
            sheet.write_string(row, 4, 'Osun State Government')
            sheet.write_string(row, 5, 'DEDUCTION ' + scenario_objs[0].payroll_id.calendar_id.name)
            row += 1   
            #Tax deductions    
            sheet.write_number(row, 0, row)
            sheet.write_string(row, 1, '9000000001')
            sheet.write_string(row, 2, '076')
            sheet.write_number(row, 3, scenario_objs[0].payroll_id.total_tax_payroll, money_format)
            sheet.write_string(row, 4, 'Osun State Government')
            sheet.write_string(row, 5, 'PAYE ' + scenario_objs[0].payroll_id.calendar_id.name)
                
            workbook.close()            
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #scenario_objs[0].update({'deduction_report': xlsx_data})
        return xlsx_data

class payment_exec_summary_report(ReportXlsx):
            
    def generate_xlsx_report(self, workbook, vals, scenario_objs, output):
        header_font = workbook.add_format({'font_size': 18, 'bold': True, 'bottom': 2, 'top': 2})
        header_font2 = workbook.add_format({'font_size': 14, 'bold': True, 'align': 'center'})
        bold_font = workbook.add_format({'bold': True})
        red_font = workbook.add_format({'font_color': 'red'})
        money_format_string = {'num_format': '###,###,##0.#0'}
        money_format = workbook.add_format(money_format_string)
        money_format_bold = workbook.add_format({'num_format': '###,###,##0.#0','bold': True})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            _logger.info("Preparing report payment_exec_summary_report...")
            #TODO Filter item_list based on scenario MDA parameter

            sheet = workbook.add_worksheet(scenario_objs[0].name[:31])
            sheet.insert_image('A1:E5', '/odoo/odoo9/osun_ippms.png')
            sheet.merge_range('A6:E7', 'PAYMENT BASED ON APPROVED SCENARIO ' + scenario_objs[0].payroll_id.calendar_id.name, header_font)
            row = 8
            indices = [0,1,2,3,4]
            header = ['Serial Number','Scenario','Percentage','Gross Amount','Net Amount']
            for c in indices:
                sheet.write(row, c, header[c], bold_font)
            row += 1

            sub_total1_gross = 0
            sub_total1_net = 0
            if scenario_objs[0].scenario_item_ids:            
                for scenario_item in scenario_objs[0].scenario_item_ids:
                    net_amount = 0
                    gross_amount = 0
                    sheet.write_number(row, 0, row - 8)
                    sheet.write_string(row, 1, scenario_item.name)
                    sheet.write_string(row, 2, (str(scenario_item.percentage) + '%'))
                    if scenario_objs[0].payment_ids and scenario_item.payscheme_ids:
                        payments = scenario_objs[0].payment_ids.filtered(lambda r: r.active and r.employee_id.payscheme_id.id in scenario_item.payscheme_ids.ids and r.employee_id.level_id.step >= scenario_item.level_min and r.employee_id.level_id.step <= scenario_item.level_max)
                        
                        for p in payments:
                            gross_amount += (p.payroll_item_id.gross_income * p.percentage / 100)
                            net_amount += (p.payroll_item_id.net_income * p.percentage / 100)
                    #TODO Filter gross and net amount from payroll using payscheme, min_level and max_level
                    sheet.write_number(row, 3, gross_amount, money_format)
                    sheet.write_number(row, 4, net_amount, money_format)
                    sub_total1_gross += gross_amount
                    sub_total1_net += net_amount
                    row += 1

                #Deductions
                #payroll_objs[0].env.cr.execute("select id from ")
                
                #payroll_objs[0].env.cr.execute("select sum(amount) from ng_state_payroll_payroll_item_line where item_id in (select id from ng_state_payroll_payroll_item where name ilike '%NHF%' payroll_id=" + str(scenario_objs[0].payroll_id.id))
                #sum_nhf_deductions = payroll_objs[0].env.cr.fetchone() #NHF Deductions
    
                #payroll_objs[0].env.cr.execute("select sum(amount) from ng_state_payroll_payroll_item_line where item_id in (select id from ng_state_payroll_payroll_item where name ilike '%NHF%' payroll_id=" + str(scenario_objs[0].payroll_id.id))
                #sum_nhf_deductions = payroll_objs[0].env.cr.fetchone() #NHF Deductions
                staff_strength = 0
                _logger.info("payment_ids=%d", len(scenario_objs[0].payment_ids))
                #MDA Deductions
                nhf_mda = 0
                paye_mda = 0
                pension_mda = 0
                deduction_other = 0
                gross_mda = 0
                #TESCOM Deductions
                nhf_tescom = 0
                paye_tescom = 0
                pension_tescom = 0
                gross_tescom = 0
                #LTH Deductions
                nhf_lth = 0
                paye_lth = 0
                pension_lth = 0
                gross_lth = 0
                for payment_item in scenario_objs[0].payment_ids:
                    staff_strength += 1
                    if payment_item.active and payment_item.employee_id.department_id.orgtype_id.id in [1,2,3]:
                        gross_mda += (payment_item.payroll_item_id.gross_income * payment_item.percentage / 100)
                        paye_mda += (payment_item.payroll_item_id.paye_tax * payment_item.percentage / 100)
                        
                        for item_line in payment_item.payroll_item_id.item_line_ids:
                            if item_line.name.upper().find('NHF') > -1:
                                nhf_mda += (item_line.amount * payment_item.percentage / 100)
                            elif item_line.name.upper().find('PENSION') > -1:
                                pension_mda += (item_line.amount * payment_item.percentage / 100)
                            else:
                                deduction_other += (item_line.amount * payment_item.percentage / 100)
                    elif payment_item.active and 'TESCOM' in payment_item.employee_id.department_id.name:
                        gross_tescom += (payment_item.payroll_item_id.gross_income * payment_item.percentage / 100)
                        paye_tescom += (payment_item.payroll_item_id.paye_tax * payment_item.percentage / 100)
                        
                        for item_line in payment_item.payroll_item_id.item_line_ids:
                            if item_line.name.upper().find('NHF') > -1:
                                nhf_tescom += (item_line.amount * payment_item.percentage / 100)
                            elif item_line.name.upper().find('PENSION') > -1:
                                pension_tescom += (item_line.amount * payment_item.percentage / 100)
                            else:
                                deduction_other += (item_line.amount * payment_item.percentage / 100)
                    elif payment_item.active and 'LTH' in payment_item.employee_id.department_id.name:
                        gross_lth += (payment_item.payroll_item_id.gross_income * payment_item.percentage / 100)
                        paye_lth += (payment_item.payroll_item_id.paye_tax * payment_item.percentage / 100)
                        
                        for item_line in payment_item.payroll_item_id.item_line_ids:
                            if item_line.name.upper().find('NHF') > -1:
                                nhf_lth += (item_line.amount * payment_item.percentage / 100)
                            elif item_line.name.upper().find('PENSION') > -1:
                                pension_lth += (item_line.amount * payment_item.percentage / 100)
                            else:
                                deduction_other += (item_line.amount * payment_item.percentage / 100)
            
                processing_fee = staff_strength * 100
                redemption_bill_mda = gross_mda * 0.05
                redemption_bill_tescom = gross_tescom * 0.05
                redemption_bill_lth = gross_lth * 0.05
                sub_total2_gross = 2 * (pension_mda + pension_tescom + pension_lth) + redemption_bill_mda + redemption_bill_tescom + redemption_bill_lth + paye_mda + paye_tescom + paye_lth
                sub_total2_net = pension_mda + pension_tescom + pension_lth + redemption_bill_mda + redemption_bill_tescom + redemption_bill_lth
                grand_total_gross = sub_total1_gross + sub_total2_gross
                grand_total_net = sub_total1_net + sub_total2_net

                sheet.write_number(row, 0, row - 8)
                sheet.write_string(row, 1, 'Processing Fees', bold_font)
                sheet.write_blank(row, 2, '')
                sheet.write_number(row, 3, processing_fee, money_format_bold)
                sheet.write_number(row, 4, processing_fee, money_format_bold)
                row += 1
                
                sheet.write_blank(row, 0, '')
                sheet.write_string(row, 1, 'SUB-TOTAL I', bold_font)
                sheet.write_blank(row, 2, '')
                sheet.write_number(row, 3, sub_total1_gross, money_format_bold)
                sheet.write_number(row, 4, sub_total2_net, money_format_bold)
                row += 1
            
                sheet.merge_range('A' + str(row + 1) + ':' + 'E' + str(row + 1), 'DEDUCTIONS', header_font2)
                row += 1
                                
                sheet.write_string(row, 0, '1')
                sheet.write_string(row, 1, 'NHF')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, (nhf_mda + nhf_tescom + nhf_lth), money_format)
                row += 1
    
                sheet.write_string(row, 0, '2a')
                sheet.write_string(row, 1, 'PAYE (MDA)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, paye_mda, money_format)
                row += 1
    
                sheet.write_string(row, 0, '2b')
                sheet.write_string(row, 1, 'PAYE (TESCOM)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, paye_tescom, money_format)
                row += 1
    
                sheet.write_string(row, 0, '2c')
                sheet.write_string(row, 1, 'PAYE (LTH)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, paye_lth, money_format)
                row += 1
    
                sheet.write_string(row, 0, '3a')
                sheet.write_string(row, 1, 'Contributory Pension (MDA)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, pension_mda, money_format)
                row += 1
    
                sheet.write_string(row, 0, '3b')
                sheet.write_string(row, 1, 'Contributory Pension (TESCOM)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, pension_tescom, money_format)
                row += 1
    
                sheet.write_string(row, 0, '3c')
                sheet.write_string(row, 1, 'Contributory Pension (LTH)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, pension_lth, money_format)
                row += 1
    
                sheet.write_string(row, 0, '4')
                sheet.write_string(row, 1, 'Other Deductions')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, deduction_other, red_font)
                row += 1
    
                sheet.write_string(row, 0, '5a')
                sheet.write_string(row, 1, 'Contributory Pension - Employer (MDA)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, pension_mda, money_format)
                row += 1
    
                sheet.write_string(row, 0, '5b')
                sheet.write_string(row, 1, 'Contributory Pension - Employer (TESCOM)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, pension_tescom, money_format)
                row += 1
    
                sheet.write_string(row, 0, '5c')
                sheet.write_string(row, 1, 'Contributory Pension - Employer (LTH)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, pension_lth, money_format)
                row += 1
    
                sheet.write_string(row, 0, '6a')
                sheet.write_string(row, 1, 'Redemption Bill - 5% of Wage Bill (MDA)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, redemption_bill_mda, money_format)
                row += 1
    
                sheet.write_string(row, 0, '6b')
                sheet.write_string(row, 1, 'Redemption Bill - 5% of Wage Bill (TESCOM)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, redemption_bill_tescom, money_format)
                row += 1
    
                sheet.write_string(row, 0, '6c')
                sheet.write_string(row, 1, 'Redemption Bill - 5% of Wage Bill (LTH)')
                sheet.write_blank(row, 2, '')
                sheet.write_blank(row, 3, '')
                sheet.write_number(row, 4, redemption_bill_lth, money_format)
                row += 1
            
                sheet.write_blank(row, 0, '')
                sheet.write_string(row, 1, 'SUB-TOTAL II', bold_font)
                sheet.write_blank(row, 2, '')
                sheet.write_number(row, 3, sub_total2_gross, money_format_bold)
                sheet.write_number(row, 4, sub_total2_net, money_format_bold)
                row += 1
            
                sheet.write_blank(row, 0, '')
                sheet.write_string(row, 1, 'GRAND TOTAL', bold_font)
                sheet.write_blank(row, 2, '')
                sheet.write_number(row, 3, grand_total_gross, money_format_bold)
                sheet.write_number(row, 4, grand_total_net, money_format_bold)
                row += 1

                sheet.merge_range('A' + str(row) + ':' + 'E' + str(row), '', header_font2)
                row += 1

                sheet.merge_range('A' + str(row) + ':' + 'E' + str(row), 'NB: OTHER SALARIES ARE INCLUDED BY THE MINISTRY OF FINANCE', header_font2)
                row += 1

            workbook.close()            
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
            #scenario_objs[0].update({'exec_summary_report': xlsx_data})
            
            _logger.info("Report payment_exec_summary_report done.")

        return xlsx_data
            
class payment_mfb_report(ReportXlsx):
    
    def generate_xlsx_report(self, workbook, vals, scenario_objs, output):
        bold_font = workbook.add_format({'bold': True})
        money_format = workbook.add_format({'num_format': '###,###,##0.#0'})
        money_format_bold = workbook.add_format({'num_format': '###,###,##0.#0','bold': True})
        xlsx_data = 0
        file_name = REPORTS_DIR + self.name + '_' + str(payroll_objs[0].id) + '.xlsx'
        try:
            with open(file_name, "rb") as xlfile:
                xlsx_data = xlfile.read()
        except IOError:
            mfb_ids = scenario_objs[0].env['ng.state.payroll.mfb'].search([('active', '=', True)])
            
            for mfb_id in mfb_ids:
                sheet_name = mfb_id.name[:31]
                sheet = workbook.get_worksheet_by_name(sheet_name)
                row = {}
                sheet = workbook.add_worksheet(sheet_name)
                row[sheet_name] = 0
                indices = [0,1,2,3,4,5]
                header = ['Serial #','Employee Name','Employee Number','Organization/TCO','Payable','MFB Account #']
                for c in indices:
                    sheet.write(row[sheet_name], c, header[c], bold_font)
                row[sheet_name] = 1
                item_list = []
                if scenario_objs[0].payroll_id.do_payroll:
                    item_list = scenario_objs[0].payment_ids.filtered(lambda r: r.active and r.employee_id.mfb_id.id == mfb_id.id)
                else:
                    item_list = scenario_objs[0].payment2_ids.filtered(lambda r: r.active and r.employee_id.mfb_id.id == mfb_id.id)
    
                sheet = None
                for payment_item in item_list:
                    sheet.write_number(row[sheet_name], 0, row[sheet_name])
                    sheet.write_string(row[sheet_name], 1, payment_item['employee_id'].name_related)
                    sheet.write_string(row[sheet_name], 2, payment_item['employee_id'].employee_no)
                    if scenario_objs[0].payroll_id.do_payroll:
                        sheet.write_string(row[sheet_name], 3, payment_item['employee_id'].department_id.name)
                    else:
                        if payment_item['employee_id'].tco_id:
                            sheet.write_string(row[sheet_name], 3, payment_item['employee_id'].tco_id.name)
                        else:
                            sheet.write_blank(row[sheet_name], 3, '')
                    sheet.write_number(row[sheet_name], 4, payment_item['amount'], money_format)
                    if payment_item['employee_id'].mfb_account:
                        sheet.write_string(row[sheet_name], 5, payment_item['employee_id'].mfb_account)
                    else:
                        sheet.write_blank(row[sheet_name], 5, '')
                    
                    row[sheet_name] += 1
    
                #Sum up
                sheet.write_blank(row[sheet_name], 0, '')
                sheet.write_string(row[sheet_name], 1, mfb_id.account_no, bold_font)
                sheet.write_string(row[sheet_name], 2, mfb_id.parent_bank_id.name, bold_font)
                sheet.write_string(row[sheet_name], 3, mfb_id.name, bold_font)
                col_name = string.ascii_uppercase[4]
                sheet.write_formula(row[sheet_name], 4, '=SUM(' + col_name + '2:' + col_name + str(row[sheet_name]) + ')', money_format_bold)
                sheet.write_blank(row[sheet_name], 5, '')
                
            workbook.close()
            xlsx_data = output.getvalue()
            with open(file_name,"wb") as f:
	            f.seek(0)
	            f.write(xlsx_data)
                #scenario_objs[0].update({'mfb_report': xlsx_data})
        return xlsx_data
        
payroll_summary_rep = payroll_summary_report('report.payroll.summary.xlsx',
            'ng.state.payroll.payroll')

pension_exec_summary_rep = pension_exec_summary_report('report.pension.exec.summary.xlsx',
            'ng.state.payroll.payroll')

payroll_exec_summary_rep = payroll_exec_summary_report('report.payroll.exec.summary.xlsx',
            'ng.state.payroll.payroll')

payroll_exec_summary2_rep = payroll_exec_summary2_report('report.payroll.exec.summary2.xlsx',
            'ng.state.payroll.payroll')

payroll_paye_rep = payroll_paye_report('report.payroll.paye.xlsx',
            'ng.state.payroll.payroll')

payroll_item_rep = payroll_item_report('report.payroll.item.xlsx',
            'ng.state.payroll.payroll')

pension_item_rep = pension_item_report('report.pension.item.xlsx',
            'ng.state.payroll.payroll')

payroll_all_rep = payroll_all_report('report.payroll.all.xlsx',
            'ng.state.payroll.payroll')

payroll_tescom_rep = payroll_tescom_report('report.payroll.tescom.xlsx',
            'ng.state.payroll.payroll')

payroll_tescom_school_rep = payroll_tescom_school_report('report.payroll.tescom.school.xlsx',
            'ng.state.payroll.payroll')

payroll_leavebonus_rep = payroll_leavebonus_report('report.payroll.leavebonus.xlsx',
            'ng.state.payroll.payroll')

payroll_mda_rep = payroll_mda_report('report.payroll.mda.xlsx',
            'ng.state.payroll.payroll')

payroll_mda_deduction_rep = payroll_mda_deduction_report('report.payroll.mda.deduction.xlsx',
            'ng.state.payroll.payroll')

payroll_mda_deduction_head_rep = payroll_mda_deduction_head_report('report.payroll.mda.deduction.head.xlsx',
            'ng.state.payroll.payroll')

payroll_mda_summarized_rep = payroll_mda_summarized_report('report.payroll.mda.summarized.xlsx',
            'ng.state.payroll.payroll')

pension_mda_rep = pension_mda_report('report.pension.mda.xlsx',
            'ng.state.payroll.payroll')

pension_tco_rep = pension_tco_report('report.pension.tco.xlsx',
            'ng.state.payroll.payroll')

payment_item_rep = payment_item_report('report.payment.item.xlsx',
            'ng.state.payroll.scenario')

payment_nibbs_rep = payment_nibbs_report('report.payment.nibbs.xlsx',
            'ng.state.payroll.scenario')

payment_exec_summary_rep = payment_exec_summary_report('report.payment.exec.summary.xlsx',
            'ng.state.payroll.scenario')

deduction_nibbs_rep = deduction_nibbs_report('report.deduction.nibbs.xlsx',
            'ng.state.payroll.scenario')

payment_mfb_rep = payment_mfb_report('report.payment.mfb.xlsx',
            'ng.state.payroll.scenario')
