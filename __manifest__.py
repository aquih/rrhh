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
    'version': '3.2',
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
        'views/hr_employee_views.xml',
        'views/hr_payslip_run_views.xml',
        'views/hr_payslip_views.xml',
        'views/res_company_views.xml',

        'report/recibo.xml',
        'report/libro_salarios.xml',
        'report/report_views.xml',

        'wizard/planilla_pdf.xml',
        'wizard/planilla.xml',
        'wizard/rrhh_libro_salarios_view.xml',
        'wizard/rrhh_informe_empleador_view.xml',
        'wizard/igss.xml',
        'wizard/rrhh_informe_isr_view.xml',
        'security/ir.model.access.csv',
        'security/rrhh_security.xml',

        # 'views/hr_leave_type_views.xml',
        # 'views/hr_contract_views.xml',
        # 'views/hr_work_entry_views.xml',
        # 'wizard/cerrar_nominas.xml',
    ],
    'license': 'Other OSI approved licence',
}
