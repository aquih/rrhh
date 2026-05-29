# -*- coding: utf-8 -*-
{
    'name': "RRHH",
    'summary': """ Módulo de RRHH para Guatemala """,
    'description': """
        Módulo de RRHH para Guatemala
    """,
    'author': "aquíH",
    'website': "http://www.aquih.com",
    'category': 'Uncategorized',
    'version': '3.10',
    'depends': ['base', 'hr_payroll_account', 'l10n_gt_extra', 'account_followup', 'hr_holidays', 'hr_work_entry'],
    'data': [
        'data/hr_payslip_input_type_data.xml',
        'data/hr_work_entry_type_data.xml',
        'data/hr_payroll_structure_data.xml',
        'data/hr_salary_rule_data.xml',
        'data/hr_leave_type_data.xml',
        'data/report_paperformat.xml',

        'views/rrhh_planilla_views.xml',
        'views/rrhh_prestamo_views.xml',
        'views/rrhh_recibo_views.xml',
        'views/hr_employee_views.xml',
        'views/hr_payslip_run_views.xml',
        'views/hr_payslip_views.xml',
        'views/res_company_views.xml',

        'report/recibo_templates.xml',
        'report/recibo_report_views.xml',
        'report/libro_salarios_templates.xml',
        'report/planilla_templates.xml',

        'wizard/hr_payslip_run_planilla_views.xml',
        'wizard/hr_employee_libro_salarios_views.xml',
        'wizard/hr_employee_informe_empleador_views.xml',
        'wizard/hr_payslip_archivo_igss_views.xml',
        'wizard/hr_payslip_run_informe_isr_views.xml',

        'security/ir.model.access.csv',
        'security/rrhh_security.xml',
    ],
    'license': 'Other OSI approved licence',
}
