class ExcelFormatRequest():
    '''
    This class contains methods related to modifying request suited to excel operations
    '''
    def form_style_req(self, input_loc, l_range, h_range, is_row=True):
        '''
        To form style request
        '''
        print("Inside form style req function")
        style_request = {}
        style_request.update(
            {
                'is_row':is_row,
                'input_loc':input_loc,
                'l_range':l_range,
                'h_range':h_range,
            }
        )
        print("Style req formed - %s", style_request)
        return style_request

    def form_formula_req(self, input_row=None, input_col=None, start_row=None, start_col=None, end_row=None, end_col=None, is_row=False):
        '''
        To form formula request
        '''
        print("Inside form formula req function")
        formula_request = {
            'input_row':input_row,
            'input_col':input_col,
            'start_row':start_row,
            'start_col':start_col,
            'end_row':end_row,
            'end_col':end_col,
            'is_row':is_row,
        }                
        print("Formula req formed - %s", formula_request)
        return formula_request

    def form_copy_formula_req(self, cp_sheet=None, cp_sheet_input_row=None, cp_sheet_input_col=None, cp_sheet_start_row=None,
        cp_sheet_start_col=None, cp_sheet_end_row=None, cp_sheet_end_col=False):
        '''
        To form copy formula request
        '''
        print("Inside form copy formula req function")
        formula_cp_request = {
            'cp_sheet':cp_sheet,
            'cp_sheet_input_row':cp_sheet_input_row,
            'cp_sheet_input_col':cp_sheet_input_col,
            'cp_sheet_start_row':cp_sheet_start_row,
            'cp_sheet_start_col':cp_sheet_start_col,
            'cp_sheet_end_row':cp_sheet_end_row,
            'cp_sheet_end_col':cp_sheet_end_col,
        }
        print("Copy formula req formed - %s", formula_cp_request)
        return formula_cp_request


class GenericFormatting(ExcelFormatRequest):
    '''
    Can add more generic formatting based operations in this class
    '''
    pass