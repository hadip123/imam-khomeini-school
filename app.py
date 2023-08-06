from flask import Flask, request
import openpyxl

app = Flask(__name__)

@app.route('/update', methods=['POST'])
def update_form():

    # Load the workbook
    wb = openpyxl.load_workbook('form.xlsx')

    # Get the worksheet
    ws = wb['form']
 
    # Name
    ws['B7'] = request.json['name']
    # LastName
    ws['J7'] = request.json['last_name'] 
    # Father Name
    ws['M7'] = request.json['father_name']
    # Field of Study
    ws['Q7'] = request.json['field_of_study']
    # Number of Identity
    ws['X7'] = request.json['id_number']
    # (Alphabet) Number of Identity
    ws['AA7'] = request.json['id_alphabet']
    # (Int) Number of Identity
    ws['AA8'] = request.json['id_number_int']
    # Code Melli
    ws['H8'] = request.json['national_code']
    # Place of Issue
    ws['M8'] = request.json['issue_place']
    # Mother Name
    ws['S8'] = request.json['mother_name']
    # (Day) Birth
    ws['D9'] = request.json['birth_day']
    # (Month) Birth
    ws['F9'] = request.json['birth_month']
    # (Year) Birth
    ws['H9'] = request.json['birth_year']
    # Country
    ws['P9'] = request.json['country']
    # City/Village
    ws['U9'] = request.json['city']
    # Postal Code
    ws['X9'] = request.json['postal_code']
    # Religion
    ws['H10'] = request.json['religion']
    # Status of body
    ws['N10'] = request.json['body_status'] 
    # Mobile number in "Shad" system
    ws['R10'] = request.json['shad_mobile']
    # Is Left-Handed?
    ws['Z10'] = request.json['left_handed']
    # Father's education
    ws['F12'] = request.json['father_education']
    # Father's occupation
    ws['K12'] = request.json['father_occupation']
    # Father's work address
    ws['O12'] = request.json['father_work_address']
    # Father's work phone
    ws['X12'] = request.json['father_work_phone']
    # (Day) Father's Birth
    ws['F13'] = request.json['father_birth_day']
    # (Month) Father's Birth  
    ws['H13'] = request.json['father_birth_month']
    # (Year) Father's Birth
    ws['J13'] = request.json['father_birth_year']
    # Father's ID
    ws['L13'] = request.json['father_id']
    # Father's Place of Issue
    ws['O13'] = request.json['father_issue_place'] 
    # Father's Insurance code
    ws['S13'] = request.json['father_insurance_code']
    # Father's National Code
    ws['X13'] = request.json['father_national_code']
    # Mother's education
    ws['F14'] = request.json['mother_education']
    # Mother's occupation
    ws['K14'] = request.json['mother_occupation']
    # Mother's work address
    ws['O14'] = request.json['mother_work_address']
    # Mother's work phone
    ws['X14'] = request.json['mother_work_phone']
    # (Day) Mother's Birth
    ws['F15'] = request.json['mother_birth_day']
    # (Month) Mother's Birth
    ws['H15'] = request.json['mother_birth_month']
    # (Year) Mother's Birth
    ws['J15'] = request.json['mother_birth_year']  

    # (Day) Supervisor's Birth
    ws['F21'] = request.json['supervisor_birth_day']
    # (Month) Supervisor's Birth
    ws['H21'] = request.json['supervisor_birth_month']
    # (Year) Supervisor's Birth
    ws['J21'] = request.json['supervisor_birth_year']


    # Mother's ID
    ws['L15'] = request.json['mother_id']
    # Mother's Place of Issue
    ws['O15'] = request.json['mother_issue_place']
    # Mother's Insurance code
    ws['S15'] = request.json['mother_insurance_code']
    # Mother's National Code
    ws['X15'] = request.json['mother_national_code']
    # Address
    ws['H16'] = request.json['address']
    # Home Telephone
    ws['R16'] = request.json['home_phone']
    # Status of housing
    ws['Y16'] = request.json['housing_status']
    # Father's phone number
    ws['E17'] = request.json['father_phone']
    # Mother's phone number
    ws['L17'] = request.json['mother_phone']  
    # Phone number
    ws['Q17'] = request.json['phone']
    # Emergency phone number
    ws['W17'] = request.json['emergency_phone']
    # Who does the student live with at home?
    ws['J18'] = request.json['live_with']
    # The student's housing situation if he lives away from his family to study: 
    ws['T18'] = request.json['housing_situation']
    # Does the student have an independent study room:
    ws['AA18'] = request.json['study_room']
    # the number of family members:
    ws['E19'] = request.json['family_members']
    # How many children are there before him?
    ws['K19'] = request.json['children_before']
    # Who is the student supervisor?
    ws['Q19'] = request.json['student_supervisor']
    # Email
    ws['W19'] = request.json['email']
    # Height  
    ws['D20'] = request.json['height']
    # Weight
    ws['K20'] = request.json['weight']
    # Ability, skill and position or rank:
    ws['P20'] = request.json['ability']
    # The number used in the government portal
    ws['P21'] = request.json['gov_portal_number']
    # Status of pervious year
    ws['H23'] = request.json['previous_year_status']
    # The total GPA of the ninth grade:
    ws['P23'] = request.json['ninth_grade_gpa']
    # The GPA of ninth grade second period
    ws['U23'] = request.json['ninth_grade_second_gpa']
    # Accepted in "Nemoone-Dolati" Test
    ws['AA23'] = request.json['dolati_test']
    # Pervious school name
    ws['H24'] = request.json['previous_school_name']  
    # Pervious school code
    ws['P24'] = request.json['previous_school_code']
    # Witness quota, sacrifice:
    ws['V24'] = request.json['witness_quota']
    # Under the cover of the Imam Khomeini (RA) relief committee:
    ws['J25'] = request.json['imam_relief']
    # Under welfare:
    ws['O25'] = request.json['welfare']
    # Does father work in Ministry of Education?
    ws['T25'] = request.json['father_education_ministry']
    # Does mother work in Ministry of Education?
    ws['Y25'] = request.json['mother_education_ministry']

    # Save workbook
    wb.save(request.json['national_code'] + '-' + request.json['name'] + '-' + request.json['last_name'] + '-' + request.json['field_of_study'] + '.xlsx')

    return "json updated!"