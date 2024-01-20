from collections import Counter
import spacy
from fuzzywuzzy import fuzz
import fitz
import os
import re
from pdfminer.high_level import extract_text
nlp = spacy.load("en_core_web_md")  # Load the model outside functions to avoid repeated loading
 
def header_dictionary():
    header_dictionary = {
    'ABSTRACT':['ABSTRACT', 'Abstract', 'Summary', 'Executive Summary'],
    
    'Introduction': ['Introduction', 'Intro', 'Overview', 'Background', 'Context'],
    
    'Literature Review': [
        'Literature Review', 'Literature Survey', 'INTERNAL FACTORS', 'EXTERNAL FACTORS',
        'Related Work', 'Review of the Literature', 'Research philosophies',
        'Literature Analysis', 'Previous work',
        'BACKGROUND: NEURAL MACHINE TRANSLATION',
        'Background: Recurrent Neural Network', 'Stage One: Educational research without a scientific research model',
        'Platos ethics of lourishing', 'PREFACE', 'REVIEW OF LITERATURE', 'Significance of the Study', 'THEORETICAL STANCE & LITERATURE REVIEW',
        
    ],
    
    'Problem Statement': ['Problem Statement', 'Statement of the Problem', 'Problem Formulation', 'LEARNING TO ALIGN AND TRANSLATE',
        'PHILOSOPHICAL APPROACHES IN RESEARCH', 'Objectives', 'CONCEPTIONS OF WHITE COLLAR CRIME', 'Essence of', 'Purpose',
    ],
    
    'Methodology': [
        'Methodology', 'Methods', 'Approach', 'Procedure', 
        'Experimental Design', 'Research Design', 'Materials and Methods',
        'Study Design', 'Methodological Approach', 'Qualitative and quantitative paradigms',
        'EXPERIMENT SETTINGS', 'Our Model', 'Description',
        'Experiments', 'Approach', 'MODELS', 'Our Mode', 'USE OF PHILOSOPHY SCIENTIFIC RESEARCH',
        'Activities', 'approach', 'Measurement', 'Procedure', 'Theory Method',
        'DATASET, VARIABLES, AND METHOD ',
    ],

    'Data Collection': [
        'Data Collection', 'Data Gathering', 'Data Acquisition',
        'Data Collection Methods', 'Data Capture' , 'Dataset',
        'Data Retrieval', 'Corpus and Preprocessing', 'Participants',
        'Experiments Setting', 'Tasks and Datasets', 'Participants',
        'Experimental Setup', 'DATASETS',  'Second Stage: Research applied to practice', 
    ],

    'Data Analysis': [
        'Our Proposed Method', 'Overview', 'Models', 'Instrumentation',
        'Data Analysis', 'Data Processing', 'Analysis Methods',
        'Statistical Analysis', 'Data Interpretation', 
        'Analysis of Data', 'Analytical Methods', 'Experiments',
        'Evaluation', 'Data Annotation', 'Dataset Analysis',
        'Evaluation Metrics', 'Runtime Analytics', 'Optimization',
        ' Innovation and Application',
    ],

    'Discussion': [
        'Discussion', 'Analysis', 'Interpretation', 
        'Comparison', 'Discussion of Results', 
        'Results Discussion', 'Results Interpretation',
        'Compared Methods', 'Conclusions', 'Optimization',
        'Experiments', 'Generated Descriptions: Fulframe evaluation',
        'Generated Descriptions: Region evaluation', 'Limitations',
        'DISCUSSION AND PRACTICAL IMPLEMENTATIONS',    
    ],
    
    'Unique Sections':[
        'CNNLM: Optimizing Sentence Modeling',
        'The Proposed Optimization Algorithm', 'Learning to align visual and language data',
        'Alignment objective', 'Decoding text segment alignments to images', 'Ontology', 'Epistemology',
        'Multimodal Recurrent Neural Network for generating descriptions', 'Prospects for Whiteheadian Thought in Biology',
        'Training Data for Summarization', 'Document Reader', 'NEURAL INTRA-ATTENTION MODE',
        'TOKEN GENERATION AND POINTER', 'HYBRID LEARNING OBJECTIVE', 'APPROACH OF PDFFIGURES','Caption Detection',
        'Region Identification', 'Figure Assignment', 'Figure Extraction' , 'SECTION TITLE EXTRACTION',
        'CNNLM: Optimizing Sentence Modeling', 'Unsupervised Training', 'Challenges of Objective Function',
        'DivSelect: Optimizing Sentence Selection', 'Objective Function', 'Diminishing Returns Property of Q(C)',
        'The Proposed Optimization Algorithm', 'Algorithm Description', 'Compared Methods', 'The Role of Educational Reform in Technology Development',
        'ENCODER: BIDIRECTIONAL RNN FOR ANNOTATING SEQUENCES', 'NEURAL NETWORKS FOR MACHINE TRANSLATION',
        'Learning to align visual and language data', 'Representing images', 'Representing sentences', 'Alignment objective',
        'Decoding text segment alignments to images', 'Multimodal Recurrent Neural Network for generating descriptions',
        'Generated Descriptions: Region evaluation', 'Gated Recurrent Neural Networks', 'Gated Recurrent Unit', 'Level of Health Condition',
        'Effect of co curricular achievements','Factors of the surrounding environment',
        
        
        
    ],
    
    'Results': [
        'Results', 'Findings', 'Outcome', 'Observations',
        'Experimental Results', 'Research Results', 
        'Study Findings', 'Experimental Results and Analysis',
        'Quantitative Results', 'Qualitative Results', 'Alignment',
        'Results and Analysis', 'Benchmark Evaluation', 
        'Results and Discussion', 'CONCLUSIONS AND IMPLICATIONS',
        'Outputs', 'CONCLUSION AND RECOMMENDATION', 'ECONOMETRIC RESULTS',
        'Discussion Recommendation and Conclusions',
    ],
    
    'Validation': [
        'Validation', 'Validation Process', 'Validation Results',
        'Experimental Validation', 'Verification and Validation'
    ],

    'Recommendations': [
        'Concluding Remarks', 'Recommendations', 'Suggestions', 'Implications',
        'Recommendations for Future Research', 
        'Suggested Actions', 'Impacts',
    ],

    'Conclusion': [
        'Conclusion', 'Summary', 'Closing Remarks', 
        'Final Thoughts', 'Conclusions and Future Work',
        'CONCLUDING THOUGHTS',               
    ], 

    'Conclusion and Future Work': [
        'Conclusion and Future Work', 
        'Conclusions and Future Directions',
        'Final Remarks and Future Work', 
        'Implications and Future Research',
        'Future Research Priorities ',
    ],
    # Add more header types and associated words/phrases
    }
    return header_dictionary

def read_pdf(filename):
    # Function to read the content of a PDF file
    text = extract_text(filename)
    
    # Remove non-ASCII characters
    text = text.encode('ascii', errors='ignore').decode()

    # Remove '(cid:xx)' patterns
    text = re.sub(r'\(cid:\d+\)', '', text)
    
    # Extract title
    title_match = re.search(r'^[^\n]+', text, re.MULTILINE)
    title = title_match.group().strip() if title_match else "Untitled"
    
    return title, text

def identify_headers_with_spacy(pdf_text, header_dictionary):
    # Load spaCy and create a PhraseMatcher
    nlp = spacy.load("en_core_web_lg")
    matcher = PhraseMatcher(nlp.vocab)

    # Add header patterns to the matcher
    for header_type, header_words in header_dictionary.items():
        header_patterns = [nlp(phrase) for phrase in header_words]
        matcher.add(header_type, None, *header_patterns)

    # Process the PDF text with spaCy
    doc = nlp(pdf_text)

    # Find matches using the PhraseMatcher
    matches = matcher(doc)

    # Extract matched headers
    matched_headers = []
    for match_id, start, end in matches:
        matched_headers.append((start, end, doc[start:end].text, header_type))

    return matched_headers

# GETTING TITLE
def count_capitals(text):
        return sum(1 for char in text if char.islower())
    
def extract_title_from_first_page(pdf_path):
    doc = fitz.open(pdf_path)
    
    # Loop through the lines on the first page
    first_page_text = doc[0].get_text("text")
    lines = first_page_text.split('\n')
    # Extract title based on criteria: at least two words and the largest font size
    title = "Untitled"
    largest_font_size = 0
    none_title = True
    next_line_same_as_title = False
    extra_title = ''
    for i, line in enumerate(lines):
        
        # Check if the line has at least two words
        if len(line.split()) >= 1:
            
            # Remove non-ASCII characters
            cleaned_line = line.encode('ascii', errors='ignore').decode()

            # Remove '(cid:xx)' patterns
            cleaned_line = re.sub(r'\(cid:\d+\)', '', cleaned_line)

            if re.match(r'^\s*arXiv:\d+\.\d+v\d+\s*\[\w+\.\w+\]\s*\d+\s+\w+\s+\d+', cleaned_line):
                #print("cleaned")
                continue
                
            # Retrieve font size for the current cleaned line from the document
            current_font_size = get_font_size_from_fitz_document(doc, cleaned_line)
            #print(current_font_size, " | ", cleaned_line)
            
            if current_font_size is not None and current_font_size < 9:
                continue
            if none_title and current_font_size is None:
                extra_title =  cleaned_line.strip()
                none_title = False
                
            if cleaned_line.count(',') > 1 or cleaned_line.count('\t') > 1:
                break
            if title != "Untitled" and cleaned_line.count(',') > 1 and ' and ' in cleaned_line.lower():
                break 
            if current_font_size is not None and current_font_size > largest_font_size:
                # Update title if the current line has the largest font size so far
                title = cleaned_line.strip()
                largest_font_size = current_font_size
                # Reset the flag for the next line being the same as the title    
                
            elif current_font_size == largest_font_size and cleaned_line.count(',') <= 1 and cleaned_line.count('\t') <= 1:
                # Append the next line to the title if it matches and contains fewer than two commas
                title += '\n' + cleaned_line.strip()
                
            elif i > 0 and cleaned_line.strip() == lines[i - 1].strip() and current_font_size is not None:
                # If the previous line is the same as the title, add it to the title
                title = cleaned_line.strip() + '\n' + title
    if title == "Untitled":
        title = extra_title
    if largest_font_size < 10:
        if extra_title and count_capitals(extra_title) < count_capitals(title):
            title = extra_title
    
    doc.close()
    return title, largest_font_size

def token_similarity(doc1, doc2, threshold=0.8):
    if not doc1.has_vector or not doc2.has_vector:
        return False
    return doc1.similarity(doc2) >= threshold

def fuzzy_string_match(str1, str2, threshold=80):
    return fuzz.ratio(str1.lower(), str2.lower()) >= threshold

def is_potential_header(previous_line, line, next_line, word_threshold):
    current_line_words = len(line.split(' ')) < word_threshold or len(line) < 50

    # Check if the previous line is empty or has fewer words than the threshold
    previous_line_words = not previous_line or len(previous_line.split()) < word_threshold

    # Check if the next line is empty or has equal to or more words than the threshold
    #next_line_words = not next_line or len(next_line.split()) >= word_threshold

    # Return True if all conditions are met, indicating a potential header

    return current_line_words and previous_line_words 

def process_text(lines, word_threshold):
    headers = []
    for i in range(1, len(lines) - 1):
        if lines[i] == '' or len(lines[i]) < 4:
            continue
        
        previous = lines[i - 1]
        next = lines[i + 1]
        
        if is_potential_header(previous, lines[i], next, word_threshold):
            headers.append(lines[i])
    return headers

def extract_font_size(span):
    return span.get('size', 9)

def get_font_size_from_fitz_document(document, line):
    # Filter out non-alphabetic characters from the line for comparison
    filtered_line = ''.join([char for char in line if char.isalpha()])

    for page in document:  # Iterate through each page in the document
        blocks = page.get_text("dict")["blocks"]  # Get text blocks from the page
        for block in blocks:
            if "lines" in block:
                for b_line in block["lines"]:
                    for span in b_line["spans"]:
                        # Filter out non-alphabetic characters from the span text
                        filtered_span_text = ''.join([char for char in span["text"] if char.isalpha()])

                        if filtered_line.strip() in filtered_span_text.strip():
                            return span["size"]  # Return the font size if the line matches
    return None  # Return None if the line is not found

def larger_font_line(document, page_num):
    page = document[page_num]
    text_blocks = page.get_text("dict")["blocks"]
    for block in text_blocks:
        if "lines" in block:
            lines = block["lines"]
            for i in range(1, len(lines) - 1):
                current, previous, next = lines[i], lines[i - 1], lines[i + 1]
                current_size = extract_font_size(current['spans'][0])
                prev_size = extract_font_size(previous['spans'][0])
                next_size = extract_font_size(next['spans'][0])

                if current_size > prev_size and current_size > next_size:
                    return True
    return False

def match_headers(pdf_text, header_dict, document, title, title_font):
    matched_headers = []
    extra_headers = []
    extra = []
    extra_fonts = []
    fonts = []
    second_run = 0
    intro_check = -10
    headers = process_text(pdf_text.split('\n'), word_threshold=7)
    header_docs = {header_type: [nlp(word) for word in words] for header_type, words in header_dict.items()}
    largest_font_size = 0
    start = False
    for line in headers:
        
        line_doc = nlp(line)
        matching_header = None
        readable_font = False
        for header_type, words in header_docs.items():
            for word in words:
                # first alphbet should be caps
                first_alphabet_index = next((i for i, c in enumerate(line) if c.isalpha()), None)
                if first_alphabet_index is not None and line[first_alphabet_index].isupper():
                    if fuzzy_string_match(word.text, line) or token_similarity(word, line_doc):
                        if header_type == "ABSTRACT" or header_type == "Introduction":
                            start = True
                        if start:
                            # Retrieve font size for the current line from the document
                            current_font_size = get_font_size_from_fitz_document(document, line)
                            if title and re.search(r'\b{}\b'.format(re.escape(line)), title) and current_font_size == title_font:
                                continue
                            if "keywords"in line.lower():
                                    continue
                            print(line, ': ', current_font_size)

                            if current_font_size is None:
                                extra_headers.append((header_type, line))
                                extra_fonts.append(current_font_size)
                                if header_type == "Introduction":
                                    intro_check = current_font_size   
                                break

                            elif current_font_size >= largest_font_size:
                                readable_font = True
                                
                                if "abstract" not in line.lower():
                                    largest_font_size = current_font_size
                                matching_header = (header_type, line)
                                
                                if header_type == "Introduction":
                                    intro_check = current_font_size
                                break
                        

        if matching_header and readable_font == True:
            matched_headers.append(matching_header)
            fonts.append(current_font_size)
    print(matched_headers, extra_headers)
    print(len(matched_headers), len(extra_headers))
    pick = False
    if matched_headers:
        for type, header in matched_headers:
            if type == "ABSTRACT" or type == "Introduction":
               pick = True
    if not matched_headers or pick == False:
        if len(matched_headers) < len(extra_headers):
            matched_headers = extra_headers
            fonts = extra_fonts
            second_run = 1
    

    unique_matched_headers = []
    final_fonts = []
    seen_headers = set()
    i = 0

    for type, header in matched_headers:
        if header not in seen_headers:
            final_fonts.append(fonts[i])
            unique_matched_headers.append(header)
            seen_headers.add(header)
        i = i + 1
    
    if second_run == 0:
        contains_none = any(item is None for item in final_fonts)
        contains_number = any(isinstance(item, (int, float)) for item in final_fonts)
        header_font = None
        if contains_none and contains_number:
            if intro_check is None or intro_check > 0:
                header_font = intro_check
        else:
            counter = Counter(final_fonts)
            header_font =  counter.most_common(1)[0][0] if counter else None

        print(header_font)
        
        compulory_headers = []
        for word in unique_matched_headers:
            compulory_headers.append(word)
        matched_headers = []
        fonts = []
        next_line_follow_header = False
        start = 0
        #print(compulory_headers)
        
        if header_font == None:
            for line in headers:
                current_font_size = get_font_size_from_fitz_document(document, line)
                if "keywords"in line.lower():
                    continue
                if  current_font_size == header_font:
                    line_doc = nlp(line)
                    for header_type, words in header_docs.items():
                        for word in words:
                            first_alphabet_index = next((i for i, c in enumerate(line) if c.isalpha()), None)
                            if first_alphabet_index is not None and line[first_alphabet_index].isupper():
                                if fuzzy_string_match(word.text, line) or token_similarity(word, line_doc):
                                    if header_type == "ABSTRACT" or header_type == "Introduction":
                                        start = 1
                                    if start == 1:
                                        match_headers = ("None", line)
                                        matched_headers.append(match_headers)
                                        fonts.append(current_font_size)
        else:
            for line in headers:
                if next_line_follow_header == True:
                    next_line_follow_header = False
                    continue
                current_font_size = get_font_size_from_fitz_document(document, line)
                if current_font_size == None:
                    continue
                if "keywords"in line.lower():
                    continue
                if line in compulory_headers or current_font_size == header_font:
                    line_doc = nlp(line)
                    for header_type, words in header_docs.items():
                        for word in words:
                            first_alphabet_index = next((i for i, c in enumerate(line) if c.isalpha()), None)
                            if first_alphabet_index is not None and line[first_alphabet_index].isupper():
                                if line in compulory_headers or fuzzy_string_match(word.text, line) or token_similarity(word, line_doc):
                                    if header_type == "ABSTRACT" or header_type == "Introduction":
                                        start = 1
                                    if start == 1:
                                        match_headers = ("None", line)
                                        matched_headers.append(match_headers)
                                        fonts.append(current_font_size)
                                        next_line_follow_header = True
                
            
        unique_matched_headers = []
        final_fonts = []
        seen_headers = set()
        i = 0
        
        for header in matched_headers:
            if header not in seen_headers:
                final_fonts.append(fonts[i])
                unique_matched_headers.append(header)
                seen_headers.add(header)
            i = i + 1
            
    print(unique_matched_headers)
    print(final_fonts)
    return list(unique_matched_headers), final_fonts

def extract_section_data_with_fonts(pdf_path, header_details, title):
    text = extract_text(pdf_path)
    cleaned_text = re.sub(r'[^\x00-\x7F]+', '', text)
    text_encoded = cleaned_text.encode('utf-8')
    section_data_list = []  # List to store section data

    for i in range(len(header_details) - 1):
        current_header, next_header = header_details[i], header_details[i + 1]
        start_index = text_encoded.find(current_header[1].encode('utf-8')) + len(current_header[1])
        end_index = text_encoded.find(next_header[1].encode('utf-8')) if i < len(header_details) - 1 else len(text_encoded)

        if start_index != -1 and end_index != -1:
            section_data = text_encoded[start_index:end_index].decode('utf-8').strip()
            section_data_list.append((current_header[0], section_data, current_header[1], current_header[2]))

    return section_data_list

def process_folder(pdf_path):
    
    title, pdf_text = read_pdf(pdf_path)
    header_dict = header_dictionary()
    
    # Open the PDF document once
    document = fitz.open(pdf_path)
    title, title_font = extract_title_from_first_page(pdf_path)
    matched_headers, header_fonts = match_headers(pdf_text, header_dict, document, title, title_font)
    
    headers = []
    header_details_font = []
    #print(matched_headers)
    #print(header_fonts)
    
    print(f"Title: {title}, PDF: {pdf_path}")
    i = 0
    for  header_type, word in matched_headers:
        font = header_fonts[i]
        headers.append(word)
        header_details_font.append((header_type, word, font))
        i= i + 1


    full_text = ""
    
    for page_num in range(len(document)):
        page = document[page_num]
        full_text += page.get_text("text")
        text_blocks = page.get_text("dict")["blocks"]
    
    section_data_list = extract_section_data_with_fonts(pdf_path, header_details_font, title)
    sections = []
    for header_type, section_data, header_text, font_size in section_data_list:
        sections.append(section_data)
    document.close()
    
    return sections, headers, title
    

# Specify the folder path containing the PDFs
folder_path = "Research_papers/1604.02748.pdf"
# Process the entire folder
section_data_list, headers, title = process_folder(folder_path)
'''
for i in range(len(section_data_list)):
    print(headers[i])
    print(section_data_list[i])
    print("\n-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-\n")
'''