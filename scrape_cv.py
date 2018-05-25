"""Converts the ordina CV to json format."""
import sys
import os
from subprocess import Popen, PIPE
import argparse
import re
import docx2txt
import json


def scrape_cv() -> None:
    """Convert ordina cv to json"""

    # Get the input, output files from the arguments
    files_dict = get_arguments()

    for ind in range(len(files_dict['input'])):
        # Convert the docx cv to a dict
        cv_dict = docx_to_dict(files_dict['input'][ind], 
                               files_dict['images'][ind])

        # Print dict to screen
        #dict_to_screen(cv_dict)

        # Write dict to json
        with open(files_dict['output'][ind], 'w') as f:
            json.dump(cv_dict, f, indent=4)
    

def get_arguments() -> dict:
    """Gets input file or directory from the command line argument, 
    and creates lists of input files, output files and image directories"""

    parser = argparse.ArgumentParser()
    parser.add_argument('input',
                        help='input docx file or directory')
    args = parser.parse_args()

    # Check input
    is_file = os.path.isfile(args.input)
    if (not is_file and not os.path.isdir(args.input)) or \
        (is_file and not args.input.lower().endswith('.docx')):
        print("Input is not an existing docx file or directory, exiting.")
        sys.exit()

    # Create list of input and output files
    list_input = []
    list_output = []
    list_images = []

    if is_file:
        list_input.append(args.input)
        dir, base = os.path.split(args.input)
        list_output.append(
            os.path.join(dir, 
                            base.rstrip('.docx')+'.json'))
        list_images.append(
            os.path.join(dir, 
                            'img_'+base.rstrip('.docx')))
        Popen("mkdir " + list_images[-1], shell=True) 
    else:
        for f in os.listdir(args.input):
            if os.path.isfile(os.path.join(args.input, f)) and f.lower().endswith('.docx') \
                and not '~$' in f:
                list_input.append(os.path.join(args.input, f))
                list_output.append(os.path.join(args.input, f.rstrip('.docx')+'.json'))
                list_images.append(os.path.join(args.input, 'img_'+f.rstrip('.docx')))
                Popen('mkdir "' + list_images[-1] + '"', shell=True) 

    return dict([('input', list_input), ('output', list_output), ('images', list_images)])


def dict_to_screen(cv_dict: dict) -> None:
    for item, val in cv_dict.items():
            print(item)
            if isinstance(val, list):
                for v in val:
                    print(repr(v))
                print('')
            else:
                print(repr(val), '\n')


def docx_to_dict(input_file: str, images_dir: str) -> dict:
    """Convert ordina cv to python dictionary"""

    # Scrape docx file
    text = docx2txt.process(input_file, images_dir)

    # Get regular expressions
    regexp_dict = compile_regexp()

    # Match global regular expression to the text
    matchObject = regexp_dict['global'].match(text)

    # Map fields to found matches
    fields = ('Name', 'Role', 'Profile', 'Motivation', 'Recommendation', 
              'Reference projects', 'Personal details', 'Diploma', 
              'Trainings', 'Certificates', 'Competences', 'Experience')
    cv_dict = dict(zip(fields, matchObject.groups()))

    # Split, cleanup and match subsections
    cv_dict['Role'] = cv_dict['Role'].split('\n\n')
    cv_dict['Profile'] = cv_dict['Profile'].replace('\n\n','\n').strip('\n')
    cv_dict['Motivation'] = cv_dict['Motivation'].replace('\n\n','\n').strip('\n')
    if cv_dict['Recommendation']:
        cv_dict['Recommendation'] = cv_dict['Recommendation'].replace('\n\n','\n').strip('\n')

    project_fields = ('Name', 'Function', 'Period')
    cv_dict['Reference projects'] = [ dict(zip(project_fields, lst.split('\n\n'))) \
                                           for lst in cv_dict['Reference projects'] \
                                               .lstrip('Name\n\nFunction\n\nPeriod\n\n\n\n') \
                                               .strip('\n').split('\n\n\n\n') ]

    personal_fields = ('Gender', 'Birth date', 'Residence')
    cv_dict['Personal details'] = dict(zip(personal_fields, cv_dict['Personal details'] \
                                           .lstrip('Gender\n\nBirth date\n\nResidence\n\n') \
                                           .strip('\n').split('\n\n') ))

    diploma_fields = ('Name', 'Institute', 'Graduation year')
    cv_dict['Diploma'] = [ dict(zip(diploma_fields, lst.split('\n\n'))) \
                                           for lst in cv_dict['Diploma'] \
                                               .lstrip('Name\n\nInstitute\n\nGraduation year\n\n\n\n') \
                                               .strip('\n').split('\n\n\n\n') ]

    training_fields = ('Name', 'Institute', 'Period')
    cv_dict['Trainings'] = [ dict(zip(training_fields, lst.split('\n\n'))) \
                                           for lst in cv_dict['Trainings'] \
                                               .lstrip('Name\n\nInstitute\n\nPeriod\n\n\n\n') \
                                               .strip('\n').split('\n\n\n\n') ]

    cv_dict['Certificates'] = cv_dict['Certificates'].strip('\n').split('\n\n\n\n\n\n')

    # Competences: Use regexp to deal with optional fields
    competences_fields = ('Functional skills', 'Soft skills', 'Technical  skills', 'Language skills')
    cv_dict['Competences'] = dict(zip(competences_fields, regexp_dict['competences'].match(cv_dict['Competences']).groups()))

    skills_fields = ('Skill', 'Months experience')
    for i in range(3):
        # split on exactly 2 newlines first
        if cv_dict['Competences'][competences_fields[i]]:
            cv_dict['Competences'][competences_fields[i]] = [ dict(zip(skills_fields, lst.split('\n\n\n\n'))) for lst in \
                re.split('(?<!\n)\n{2}(?!\n)', cv_dict['Competences'][competences_fields[i]] \
                  .strip('\n')) ]

    cv_dict['Competences']['Language skills'] = cv_dict['Competences']['Language skills'].strip('\n') \
            .rstrip('\n\n\n\nWriting\n\n\n\n\n\nListening\n\n\n\n\n\nSpeaking') \
            .split('\n\n\n\nWriting\n\n\n\n\n\nListening\n\n\n\n\n\nSpeaking\n\n\n\n\n\n')

    # Experience
    cv_dict['Experience'] = cv_dict['Experience'].lstrip('Period\n\n').strip('\n').split('\n\n\n\nPeriod\n\n')

    experience_fields = ('Period', 'Employer', 'Organization', 'Project name', 
                         'Function title', 'Branch', 'Situation', 'Tasks', 
                         'Result', 'Functional skills', 'Soft skills', 'Technical skills')
    for ind, exp in enumerate(cv_dict['Experience']):
        cv_dict['Experience'][ind] = dict(zip(experience_fields, regexp_dict['experience'].match(exp).groups()))
        cv_dict['Experience'][ind]['Situation'] = cv_dict['Experience'][ind]['Situation'].replace('\n\n', '\n')
        cv_dict['Experience'][ind]['Tasks'] = cv_dict['Experience'][ind]['Tasks'].replace('\n\n', '\n')
        if cv_dict['Experience'][ind]['Result']:
            cv_dict['Experience'][ind]['Result'] = cv_dict['Experience'][ind]['Result'].replace('\n\n', '\n').strip('\n')
        if cv_dict['Experience'][ind]['Functional skills']:
            cv_dict['Experience'][ind]['Functional skills'] = cv_dict['Experience'][ind]['Functional skills'] \
                                                                .strip(' ').strip('\n').split('\n\n\n\n')
        if cv_dict['Experience'][ind]['Soft skills']:
            cv_dict['Experience'][ind]['Soft skills'] = cv_dict['Experience'][ind]['Soft skills'] \
                                                                .strip(' ').strip('\n').split('\n\n\n\n')
        if cv_dict['Experience'][ind]['Technical skills']:
            cv_dict['Experience'][ind]['Technical skills'] = cv_dict['Experience'][ind]['Technical skills'] \
                                                                .strip(' ').strip('\n').split('\n\n\n\n')
    
    return cv_dict


def compile_regexp() -> dict:
    regexp_dict = {}
    
    regexp_dict['global'] = re.compile(r'''
    ^                                   # Beginning of document
    Curriculum\ Vitae
    \s*                                  
    ([\s\S]+)                           # Name
    \n{6}
    Role
    \s*
    ([\s\S]+)                           # Role
    \n{4}
    Profile
    \s*
    ([\s\S]+)                           # Profile
    \n{4}
    Motivation
    \s*
    ([\s\S]+?)                          # Motivation, 
                                        # +? to not include Recommendation
    (?:\n{4}                            # Optional field (?: non-capturing group)
    Recommendation                      
    \s*
    ([\s\S]+)                           # Recommendation
    )?                                
    \n{4}
    Reference\ projects
    \s*
    ([\s\S]+)                           # Reference projects
    \n{4}
    Personal\ details
    \n\n\n\n\n\n
    Personal\ details
    \s*
    ([\s\S]+)                           # Personal details
    \n{4}
    Diploma
    \s*
    ([\s\S]+)                           # Diploma
    \n{4}
    Trainings
    \s*
    ([\s\S]+)                           # Trainings
    \n{4}
    Certificates
    \s*
    ([\s\S]+)                           # Certificates
    \n{4}
    Competences
    \s*
    ([\s\S]+)                           # Competences
    \n{4}
    Experience
    \s*
    ([\s\S]+)                           # Experience
    \n{2}
    Ordina\ Belgium,\ Blarenberglaan 
    \ 3B,\ 2800\ Mechelen\n\nTel:
    \ \+32\ 15\ 28\ 59\ 59\ –
    \ www.ordina.be
    [\s\S]+
    $                                   # End of document
    ''', re.VERBOSE)

    regexp_dict['experience'] = re.compile(r'''
    ^                                   # Beginning of string
    ([\s\S]+)                           # Period
    \n\n
    Employer                      
    \n\n
    ([\s\S]+)                           # Employer
    \n\n
    Organization
    \n\n
    ([\s\S]+)                           # Organization
    \n\n
    Project\ name
    \n\n
    ([\s\S]+)                           # Project name
    \n\n
    Function\ title
    \n\n
    ([\s\S]+?)                          # Function title
                                        # +? to not include optional Branch
    \n\n
    (?:                                 # Optional field (?: non-capturing group)
    Branch
    \n\n
    ([\s\S]+)                           # Branch
    \n\n
    )?
    Situation
    \n\n
    ([\s\S]+)                           # Situation
    \n\n
    Tasks,\ responsibilities
    \ and\ activities
    \n\n
    ([\s\S]+?)                          # Tasks
    (?:                                 # Optional field (?: non-capturing group)
    \n\n
    Result
    \n\n
    ([\s\S]+?)                          # Result
    )?
    (?:                                 # Optional field (?: non-capturing group)
    \n\n
    Functional\ skills
    \n\n
    ([\s\S]+?)                          # Functional skills
    )?
    (?:                                 # Optional field (?: non-capturing group)
    \n\n
    Soft\ skills
    \n\n
    ([\s\S]+?)                          # Soft skills
    )?
    (?:                                 # Optional field (?: non-capturing group)
    \n\n
    Technical\ skills
    \n\n
    ([\s\S]+?)                          # Technical skills
    )?
    $                                   # End of string
    ''', re.VERBOSE)

    regexp_dict['competences'] = re.compile(r'''
    ^                                   # Beginning of string
    Basic\n\nMedium
    \n\nHigh\n\nExpert
    \n\nMonths\ experience
    \n\n
    (?:                                 # Optional field (?: non-capturing group)
    Functional\ skills                      
    \s*
    ([\s\S]+?)                          # Functional skills, non-greedy match
    )?
    \s*
    (?:                                 # Optional field (?: non-capturing group)
    Soft\ skills                      
    \s*
    ([\s\S]+?)                          # Soft skills, non-greedy match
    )?
    \s*
    (?:                                 # Optional field (?: non-capturing group)
    Technical\ skills                      
    \s*
    ([\s\S]+?)                          # Technical  skills, non-greedy match
    )?
    \s*
    Basic\n\nIntermediate
    \n\nFluent\n\nNative
    \n\nLanguage\ skills                      
    \s*
    ([\s\S]+)                           # Language skills
    $                                   # End of string
    ''', re.VERBOSE)

    return regexp_dict

if __name__ == "__main__":

    scrape_cv()