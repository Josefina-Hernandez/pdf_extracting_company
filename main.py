import fitz

import re

import openpyxl as xl

def if_companyname(string, previouse_string):
    deviding_flg = True
    for c in string:
        if c.islower():
            return False

    if 'LTD' in string:

        for c in previouse_string:
            if c.islower():
                deviding_flg = False
                break

        if if_phone(string=previouse_string):
            deviding_flg = False

        if deviding_flg:
            return previouse_string + '' + string
        else:
            return string

    else:
        return False


def if_shortname(full_name, string):
    pattern = re.compile(r'^[A-Z]+$')
    full_name_lst = str(full_name).split(' ')
    if pattern.match(string):
        judge_dict = []
        for c in string:
            judge = False
            for word in full_name_lst:
                if c in word:
                    judge = True

            judge_dict.append(judge)

        if False in judge_dict:
            return False
        else:
            return True
    else:
        return False

def if_website(string):
    result = False
    if 'www.' in string:
        result = True
    if '.com' in string:
        result = True
    if '.co.th' in string:
        result = True
    if '.co.jp' in string:
        result = True
    if '.cn' in string:
        result = True
    if '.jp' in string:
        result = True

    if '@' in string:
        result = False
    return result

def if_products_services_first(string):
    pattern = re.compile(r'^K  ')
    if pattern.match(string):
        return True
    else:
        return False

def if_products_services(flg, string):
    if flg:
        pattern0 = re.compile(r'^G  ')
        pattern1 = re.compile(r'^K  ')
        pattern2 = re.compile(r'^1  ')
        pattern3 = re.compile(r'^V  ')
        pattern4 = re.compile(r'^L  ')
        pattern5 = re.compile(r'^B  ')
        pattern6 = re.compile(r'^O  ')
        pattern7 = re.compile(r'^P  ')
        pattern8 = re.compile(r'^X  ')

        if pattern0.match(string):
            return False
        elif pattern1.match(string):
            return False
        elif pattern2.match(string):
            return False
        elif pattern3.match(string):
            return False
        elif pattern4.match(string):
            return False
        elif pattern5.match(string):
            return False
        elif pattern6.match(string):
            return False
        elif pattern7.match(string):
            return False
        elif pattern8.match(string):
            return False
        else:
            return True
    else:
        return False


def if_executive_first(string):
    pattern = re.compile(r'^G  ')
    if pattern.match(string):
        return True
    else:
        return False

def if_executive(flg, string):
    if flg:
        pattern0 = re.compile(r'^G  ')
        pattern1 = re.compile(r'^K  ')
        pattern2 = re.compile(r'^1  ')
        pattern3 = re.compile(r'^V  ')
        pattern4 = re.compile(r'^L  ')
        pattern5 = re.compile(r'^B  ')
        pattern6 = re.compile(r'^O  ')
        pattern7 = re.compile(r'^P  ')
        pattern8 = re.compile(r'^X  ')

        if pattern0.match(string):
            return False
        elif pattern1.match(string):
            return False
        elif pattern2.match(string):
            return False
        elif pattern3.match(string):
            return False
        elif pattern4.match(string):
            return False
        elif pattern5.match(string):
            return False
        elif pattern6.match(string):
            return False
        elif pattern7.match(string):
            return False
        elif pattern8.match(string):
            return False
        else:
            return True
    else:
        return False

def if_year_staff_fund(string):
    pattern1 = re.compile(r'^1  .*P  .*B  ')
    pattern2 = re.compile(r'^1  .*B  ')
    pattern3 = re.compile(r'^1  ')

    if pattern1.match(string):
        return 1

    if pattern2.match(string):
        return 2

    if pattern3.match(string):
        return 3

    return 0

def if_boi(string):
    pattern = re.compile(r'^V  ')
    if pattern.match(string):
        return True
    else:
        return False


def if_shareholders(string):
    pattern = re.compile(r'^L  ')
    if pattern.match(string):
        return True
    else:
        return False


def if_address_first(string):
    pattern = re.compile(r'^\uf075 ')
    if pattern.match(string):
        return True
    else:
        return False


def if_address(flg, string):
    if flg:
        pattern0 = re.compile(r'^G  ')
        pattern1 = re.compile(r'^K  ')
        pattern2 = re.compile(r'^1  ')
        pattern3 = re.compile(r'^V  ')
        pattern4 = re.compile(r'^L  ')
        pattern5 = re.compile(r'^B  ')
        pattern6 = re.compile(r'^O  ')
        pattern7 = re.compile(r'^P  ')
        pattern8 = re.compile(r'^X  ')

        if pattern0.match(string):
            return False
        elif pattern1.match(string):
            return False
        elif pattern2.match(string):
            return False
        elif pattern3.match(string):
            return False
        elif pattern4.match(string):
            return False
        elif pattern5.match(string):
            return False
        elif pattern6.match(string):
            return False
        elif pattern7.match(string):
            return False
        elif pattern8.match(string):
            return False
        elif if_phone(string):
            return False
        elif if_email(string):
            return False
        else:
            return True
    else:
        return False


def if_phone(string):
    pattern = re.compile(r'^[0-9~\-/\s*,]+$')
    pattern_out = re.compile(r'[0-9]+$')
    if pattern.match(string) and not pattern_out.match(string):
        return True

    else:
        return False


def if_email(string):
    pattern = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')
    if pattern.match(string):
        return True
    else:
        return False

def reading(pageNum):
    page=f[pageNum]#[434]
    rect=page.rect
    clip=fitz.Rect(0,0, rect.width, rect.height)
    a_text=page.get_text(clip=clip)
    col_dict=a_text.split('\n')
    print(a_text)
    print(col_dict)

    page_lst=[]


    CATEGORY = col_dict[3]
    ifCompanyName = False
    CompanyName = ''

    PRODUCT_SERVICES = ''
    product_services_flg = False

    EXECUTIVE = ''
    executive_flg = False


    ADDRESS = ''
    address_flg = False

    PHONE = ''
    EMAIL = ''

    line = {
        'companyname': '',
        'shortname': '',
        'website': '',
        'category': '',
        'productservices': '',
        'executive': '',
        'foundyear': '',
        'employees': '',
        'fund': '',
        'boi': '',
        'shareholders': '',
        'address': '',
        'phone': '',
        'email': '',
        'page_num': ''
    }

    previous_each = ''
    for each in col_dict:
        each =each.strip()
        if each.strip()=='':
            previous_each = each
            continue

        #print(each)

        if ifCompanyName:
            if if_shortname(full_name=CompanyName, string=str(each)):
                line['shortname'] = str(each)

                ifCompanyName = False

                previous_each = each
                continue

        companynameOrFalse = if_companyname(string=str(each), previouse_string=previous_each)
        if companynameOrFalse:
            ifCompanyName = True
            CompanyName = companynameOrFalse
            if line != {
                'companyname': '',
                'shortname': '',
                'website': '',
                'category': '',
                'productservices': '',
                'executive': '',
                'foundyear': '',
                'employees': '',
                'fund': '',
                'boi': '',
                'shareholders': '',
                'address': '',
                'phone': '',
                'email': '',
                'page_num': ''
            }:
                line['category'] = CATEGORY
                line['page_num'] = pageNum + 1
                if line['companyname']:
                    page_lst.append(line)

                line = {
                    'companyname' : '',
                    'shortname' : '',
                    'website' : '',
                    'category': '',
                    'productservices': '',
                    'executive': '',
                    'foundyear' : '',
                    'employees' : '',
                    'fund' : '',
                    'boi' : '',
                    'shareholders' : '',
                    'address' : '',
                    'phone' : '',
                    'email' : '',
                    'page_num': ''
                }
            line['companyname'] = companynameOrFalse

            PRODUCT_SERVICES = ''
            EXECUTIVE = ''
            ADDRESS = ''
            PHONE = ''
            EMAIL = ''

            previous_each = each
            continue
        else:
            ifCompanyName = False

        if if_website(str(each)):
            line['website'] = str(each)

            ifCompanyName = False

            previous_each = each
            continue

        if if_products_services_first(str(each)):
            PRODUCT_SERVICES = str(each)
            product_services_flg = True

            ifCompanyName = False

            line['productservices'] = PRODUCT_SERVICES

            previous_each = each
            continue

        if if_products_services(flg=product_services_flg, string=str(each)):
            PRODUCT_SERVICES += '\n'+str(each)
            line['productservices'] =PRODUCT_SERVICES

            previous_each = each
            continue
        else:
            product_services_flg = False

        if if_executive_first(str(each)):
            EXECUTIVE = str(each)[3:]
            executive_flg = True

            ifCompanyName = False

            line['executive'] = EXECUTIVE

            previous_each = each
            continue

        if if_executive(flg=executive_flg, string=str(each)):
            EXECUTIVE += ' ' + str(each)
            line['executive'] =EXECUTIVE

            previous_each = each
            continue
        else:
            executive_flg = False


        isYearStaffFund = if_year_staff_fund(str(each))
        if isYearStaffFund:
            if isYearStaffFund == 1:  #r'^1  .*P  .*B  '
                fund_dist1 = str(each).split('P  ')
                line['foundyear'] = fund_dist1[0][3:]
                fund_dist2 = fund_dist1[1].split('B  ')
                line['employees'] = fund_dist2[0]
                line['fund'] = fund_dist2[1]

            elif isYearStaffFund == 2: #r'^1  .*B  '
                fund_dist1 = str(each).split('B  ')
                line['foundyear'] = fund_dist1[0][3:]
                line['fund'] = fund_dist1[1]

            else: #r'^1  '
                line['foundyear'] = str(each)[3:]

            if product_services_flg:
                product_services_flg = False

            if executive_flg:
                executive_flg = False

            ifCompanyName = False

            previous_each = each
            continue



        if if_boi(str(each)):
            line['boi'] = str(each)[3:]
            if product_services_flg:
                product_services_flg = False

            if executive_flg:
                executive_flg = False

            ifCompanyName = False

            previous_each = each
            continue

        if if_shareholders(str(each)):
            line['shareholders'] = str(each)[3:]
            if product_services_flg:
                product_services_flg = False

            if executive_flg:
                executive_flg = False

            ifCompanyName = False

            previous_each = each
            continue

        if if_address_first(str(each)):
            if ADDRESS:
                ADDRESS += '\n' + str(each).replace('\uf075 ', '')
            else:
                ADDRESS = str(each).replace('\uf075 ', '')

            line['address'] = ADDRESS
            address_flg = True

            ifCompanyName = False

            previous_each = each
            continue

        if if_address(flg=address_flg, string=str(each)):
            ADDRESS += ' ' + str(each)
            line['address'] = ADDRESS

            ifCompanyName = False

            previous_each = each
            continue
        else:
            address_flg = False

        if if_phone(str(each)):
            if PHONE:
                PHONE += '\n' + str(each)
            else:
                PHONE = str(each)
            line['phone']=PHONE

            ifCompanyName = False

            previous_each = each
            continue

        if if_email(str(each)):
            if EMAIL:
                EMAIL += '\n' + str(each)
            else:
                EMAIL = str(each)
            line['email'] = EMAIL

            ifCompanyName = False

            previous_each = each
            continue

        previous_each=each
    line['category'] = CATEGORY
    line['page_num'] = pageNum + 1
    if line['companyname']:
        page_lst.append(line)

    return page_lst


class Excel_Con():
    def __init__(self, output_path):
        self.wb = xl.Workbook()
        self.ws = self.wb.active
        self.output_path = output_path
        self.create_headers()
        self.wb.save(self.output_path)

    def create_headers(self):
        self.ws.append(['No.',
                        '企業名',
                        '略称',
                        'Website',
                        'カテゴリー',
                        '産業分野',
                        '担当者',
                        '設立年',
                        '従業員数',
                        '資本金',
                        'BOI否',
                        '国による資金比率',
                        '住所',
                        '電話番号',
                        'Email',
                        'PDFのページ数'
                        ])

    def writing_data(self, no, data_line):
        self.ws.append([
            no,
            data_line['companyname'],
            data_line['shortname'],
            data_line['website'],
            data_line['category'],
            data_line['productservices'],
            data_line['executive'],
            data_line['foundyear'],
            data_line['employees'],
            data_line['fund'],
            data_line['boi'],
            data_line['shareholders'],
            data_line['address'],
            data_line['phone'],
            data_line['email'],
            data_line['page_num']
        ])

    def save_file(self):
        self.wb.save(self.output_path)
        self.wb.close()

if __name__ == '__main__':
    ec = Excel_Con(output_path='企業年鑑情報.xlsx')

    f = fitz.open('import_pdf.pdf')
    row = 1
    data=[]
    for pageNum in range(5, 492):#(22, 23): #701
        page_lst=reading(pageNum=pageNum)
        #for dict in lst:
        #    row+=1
        #    print(dict)
        data.append(page_lst)
        print(pageNum+1,'!--------------------------------------------------------------!')

    counter = 0
    page = 0
    for each_page in data:
        #print(each_page)
        page += 1
        for each_line in each_page:
            counter += 1
            print(f'------{page}----------------------{counter}!!!!!!!!!!!!!!!', '\n', each_line)
            ec.writing_data(no=counter, data_line=each_line)
        print('**************************************************')

    ec.save_file()
    f.close()
