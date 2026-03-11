/**
 * One-time script to add educationalActivities, publications, patents columns to CSV
 */
import fs from 'fs';
import { parse } from 'csv-parse/sync';

const csvPath = 'config/textcraft-europe.csv';
const content = fs.readFileSync(csvPath, 'utf8');
const records = parse(content, { columns: true, skip_empty_lines: true, relax_quotes: true, relax_column_count: true });

// Educational activities per role/department pattern
const educationData = {
  // Executives
  'CEO': "['MBA - Bocconi University','INSEAD Executive Leadership Program']",
  'COO': "['MSc Operations Research - KTH Stockholm','MIT Sloan Executive Education']",
  'CFO': "['HEC Paris MBA','CFA Institute Chartered Financial Analyst Program']",
  'CCO': "['MA Visual Communication - UdK Berlin','London College of Communication Creative Leadership']",
  'Chief Client Officer': "['BA Business Administration - University of Oxford','IMD Business School Client Management Program']",
  // Key Account roles
  'Key Account Director': "['MBA - Trinity College Dublin','INSEAD Key Account Management Program']",
  'Key Account Manager': null, // will be generated per specialization
  // Editorial
  'Editorial Director': "['MA English Literature - University of Edinburgh','Oxford Editing & Publishing Programme']",
  'Senior Editor': null, // per specialization
  'Style Editor': null, // per specialization
  'Copy Chief': null,
  // Creative Writing
  'Academic Writer Lead': "['PhD English Literature - University of Cambridge','Harvard Writing Program Fellowship']",
  'Academic Writer': null,
  'Corporate Writer Lead': "['MA Corporate Communications - LSE','Wharton Business Writing Certificate']",
  'Corporate Writer': null,
  'Technical Writer Lead': "['MSc Computer Science - ETH Zurich','Society for Technical Communication Certification']",
  'Technical Writer': null,
  'Literary Writer Lead': "['MFA Creative Writing - Sorbonne University','Iowa Writers Workshop Fellowship']",
  'Literary Writer': null,
  'Marketing Copywriter Lead': "['MA Marketing Communications - University of Amsterdam','Google Digital Marketing Certificate']",
  'Marketing Copywriter': null,
  // Round Table
  'Round Table Director': "['MSc Project Management - ESCP Business School','Certified Facilitator - IAF']",
  'Round Table Coordinator': null,
  // QA
  'QA Director': "['MA Linguistics - University of Cambridge','ISO 17100 Translation Quality Certification']",
  'Senior Proofreader': null,
  'Proofreader': null,
  'Fact Checker': null,
  'Consistency Checker': null,
  'Plagiarism Analyst': null,
  // Operations
  'Operations Director': "['MBA - WHU Otto Beisheim School','PMP Certification']",
  'HR Manager': "['MA Human Resources - University of Amsterdam','CIPD Level 7 Advanced Diploma']",
  'Finance Manager': "['BSc Accounting - University of Manchester','ACCA Certification']",
  'IT Manager': "['MSc Information Systems - Charles University Prague','ITIL Expert Certification']",
  'IT Support Specialist': "['BSc Computer Science - University of Zurich','CompTIA A+ Certification']",
  'HR Specialist': "['BA Psychology - University of Copenhagen','SHRM-CP Certification']",
  'Accountant': "['BSc Economics - Sapienza University Rome','Italian Chartered Accountant Qualification']",
  'Office Manager': "['BA Business Administration - Universidad Complutense Madrid']",
  'Administrative Assistant': "['BA Administration - University of Warsaw','Microsoft Office Specialist Certification']",
  'Systems Administrator': "['BSc Computer Engineering - University of Bristol','AWS Solutions Architect Certification']",
};

// Specialization-based education for null roles
const specEducation = {
  'Fashion & Luxury': "['BA Fashion Marketing - Istituto Marangoni Milan','Luxury Brand Management - ESSEC']",
  'Technology': "['BSc Computer Science - Universidad Politécnica de Madrid','Salesforce Certified Administrator']",
  'Manufacturing': "['BSc Industrial Engineering - Warsaw University of Technology','APICS Supply Chain Professional']",
  'Finance': "['MSc Finance - University of Amsterdam','Dutch Financial Markets Certification']",
  'Automotive': "['BSc Mechanical Engineering - TU Munich','VDA Quality Management Certificate']",
  'Design': "['BA Industrial Design - Politecnico di Milano','UX Design Certificate - Nielsen Norman Group']",
  'Healthcare': "['MSc Health Communication - Uppsala University','European Medical Writers Association Certificate']",
  'Tourism': "['BA Tourism Management - University of Lisbon','UNWTO Tourism Excellence Certificate']",
  'Government': "['Sciences Po Paris - MA Public Affairs','EU Public Administration Certificate']",
  'Education': "['MEd Education Policy - University of Vienna','Cambridge CELTA Certification']",
  'Retail': "['BSc Business Economics - Copenhagen Business School','Retail Management Certificate']",
  'Fiction & Poetry': "['MFA Creative Writing - Sorbonne University','Alliance Française Literature Prize']",
  'Documentation': "['MA Technical Communication - University of Stuttgart','tekom Certified Technical Writer']",
  'Advertising': "['BA Communications - Sapienza University Rome','Cannes Lions School Certificate']",
  'Research Papers': "['PhD Linguistics - Lund University','Nordic Research Writing Fellowship']",
  'Business Reports': "['MA Business Communication - University of Amsterdam','CFA Society Report Writing']",
  'Translation Quality': "['MA Translation Studies - Universidad de Salamanca','DipTrans IoL Certification']",
  'Localisation': "['MA Applied Linguistics - Trinity College Dublin','GILT Localisation Certificate']",
  'Legal & Formal': "['LLM International Law - University of Vienna','Austrian Legal Translation Certificate']",
  'Global Standards': "['MA International Communications - University of Zurich','ISO Technical Writing Certification']",
  'Research & Journals': "['PhD English Literature - University of Cambridge','Royal Society of Literature Fellow']",
  'Scientific Papers': "['PhD Chemistry - Heidelberg University','European Medical Writers Association Member']",
  'Medical Research': "['PhD Clinical Research - Karolinska Institutet','AMWA Medical Writing Certificate']",
  'Social Sciences': "['PhD Sociology - Sciences Po Paris','European Sociological Association Fellow']",
  'Engineering': "['PhD Mechanical Engineering - Warsaw University of Technology','IEEE Senior Member']",
  'Environmental Studies': "['PhD Environmental Science - University of Lisbon','IEMA Environmental Management Certificate']",
  'Economics': "['PhD Economics - University of Mannheim','German Economic Association Fellow']",
  'Literature Review': "['PhD Comparative Literature - University of Edinburgh','Scottish Literary Studies Fellowship']",
  'Annual Reports': "['MBA - Freie Universität Berlin','German Association for Financial Analysis Certificate']",
  'Sustainability Reports': "['MSc Sustainability - University of Amsterdam','GRI Standards Certified Training']",
  'EU Communications': "['MA European Studies - College of Europe Bruges','EU Communications Certificate']",
  'Press Releases': "['BA Journalism - Sapienza University Rome','Italian Press Association Certificate']",
  'Internal Communications': "['MA Corporate Communication - BI Norwegian Business School','IABC Certification']",
  'Investor Relations': "['MSc Finance - Universidad Carlos III de Madrid','IR Magazine Best Practice Certificate']",
  'Executive Speeches': "['MA Rhetoric - University of Edinburgh','Toastmasters Advanced Communicator Gold']",
  'Banking & Finance': "['MSc Banking - University of Zurich','Swiss Finance Institute Certificate']",
  'Software Documentation': "['MSc Computer Science - Imperial College London','Certified Professional Technical Communicator']",
  'API Documentation': "['MSc Software Engineering - TU Berlin','OpenAPI Specification Expert']",
  'User Manuals': "['BA Technical Writing - Aarhus University','Society for Technical Communication Advanced']",
  'Hardware Guides': "['BSc Electrical Engineering - Universidad Politécnica de Madrid','EU Machinery Directive Certification']",
  'Engineering Specs': "['MSc Mechanical Engineering - Warsaw University of Technology','ISO TC 10 Technical Drawing Certification']",
  'Aerospace': "['MSc Aerospace Engineering - ISAE-SUPAERO Toulouse','ESA Technical Writing Programme']",
  'Automotive': "['BSc Automotive Engineering - University of Stuttgart','VDI Technical Documentation Certificate']",
  'Telecom': "['MSc Telecommunications - KTH Stockholm','ETSI Standards Certification']",
  'Short Stories': "['MFA Creative Writing - University of Amsterdam','European Short Story Prize Nominee']",
  'Poetry': "['MFA Poetry - Trinity College Dublin','Seamus Heaney Centre Fellowship']",
  'Dramatic Writing': "['MA Dramatic Writing - Accademia Nazionale d Arte Drammatica Rome','European Playwriting Prize']",
  'Nordic Noir': "['MA Scandinavian Studies - University of Copenhagen','Danish Writers Academy Graduate']",
  'Historical Fiction': "['MA History - University of Warsaw','Polish Literary Institute Fellowship']",
  'Contemporary Fiction': "['MFA Creative Writing - University of East Anglia','Booker Prize Longlist 2024']",
  'Magical Realism': "['MFA Creative Writing - University of Lisbon','José Saramago Foundation Fellow']",
  'Brand Campaigns': "['BA Marketing - University of Mannheim','German Advertising Federation Certificate']",
  'Digital Marketing': "['BSc Digital Marketing - University of Mannheim','Google Ads Certification']",
  'Social Media': "['BA Communications - University of Amsterdam','Meta Blueprint Certification']",
  'Video Scripts': "['BA Film Studies - Universidad Complutense Madrid','Adobe Certified Professional']",
  'Retail Campaigns': "['BA Marketing - BI Norwegian Business School','Nordic Retail Communication Award']",
  'Luxury Brands': "['BA Luxury Marketing - Bocconi University','Comité Colbert Excellence Certificate']",
  'Tourism & Travel': "['BA Tourism Communications - Heidelberg University','UNWTO Communications Certificate']",
  'Hospitality': "['BA Hospitality Management - University of Lisbon','European Hospitality Communications Certificate']",
  'Session Management': "['MSc Project Management - ESCP Paris','Certified Professional Facilitator']",
  'Cross-Dept Sessions': "['MA Organisational Psychology - Humboldt University Berlin','ICF Team Coaching Certificate']",
  'Client Reviews': "['BA Business Administration - University of Amsterdam','Client Success Manager Certificate']",
  'Creative Sessions': "['BA Fine Arts - University of Lisbon','Design Thinking Facilitator Certificate']",
  'Editorial Reviews': "['MA Publishing - Stockholm University','Swedish Publishers Association Certificate']",
  'Style Discussions': "['MA Linguistics - Sapienza University Rome','European Style Guide Certification']",
  'Technical Reviews': "['MSc Computer Science - Charles University Prague','IEEE Technical Review Board Member']",
  'Final Approvals': "['BA English - University College Dublin','ISO 9001 Quality Auditor Certification']",
  'Standards': "['MA Linguistics - University of Cambridge','ISO 17100 Translation Quality Lead Auditor']",
  'French Texts': "['MA French Language - Sorbonne University','Académie Française Proofreading Certificate']",
  'German Texts': "['MA German Philology - Ludwig Maximilian University Munich','Duden Proofreading Certification']",
  'Italian Texts': "['MA Italian Studies - University of Bologna','Accademia della Crusca Language Certificate']",
  'Spanish Texts': "['MA Hispanic Studies - Universidad Complutense Madrid','RAE Language Certification']",
  'Polish Texts': "['MA Polish Philology - University of Warsaw','Polish Language Council Certificate']",
  'Research Verification': "['MSc Information Science - Humboldt University Berlin','Certified Fact Checker - IFCN']",
  'Academic Verification': "['MSc Research Methods - Lund University','Nordic Research Integrity Network Member']",
  'Style Compliance': "['MA Portuguese Linguistics - University of Lisbon','European Style Compliance Certificate']",
  'Academic Integrity': "['MA Academic Ethics - University College Dublin','Turnitin Certified Educator']",
  'Operations': "['MBA - WHU Otto Beisheim School','PRINCE2 Practitioner']",
  'People Operations': "['MA Human Resources - University of Amsterdam','CIPD Level 7 Advanced Diploma']",
  'Budgets & Reporting': "['BSc Accounting - University of Manchester','CIMA Management Accounting Certificate']",
  'Infrastructure': "['MSc Information Systems - Charles University Prague','ITIL Expert Certification']",
  'User Support': "['BSc Computer Science - University of Zurich','ITIL Foundation Certificate']",
  'Recruitment': "['BA Psychology - University of Copenhagen','LinkedIn Recruiter Certification']",
  'Payroll': "['BSc Economics - Sapienza University Rome','Italian Payroll Professional Certificate']",
  'Facilities': "['BA Business Administration - Universidad Complutense Madrid','IFMA Facility Management Certificate']",
  'Scheduling': "['BA Administration - University of Warsaw','Microsoft Project Certification']",
  'Cloud Systems': "['BSc Computer Engineering - University of Bristol','AWS Solutions Architect Professional']",
};

// Publications per specialization
const publicationsData = {
  'Leadership': "['The Future of European Publishing (Harvard Business Review 2024)','Leading Creative Teams Across Borders']",
  'Enterprise Clients': "['Enterprise Content Strategy in the AI Age','Building Long-Term Client Partnerships in Publishing']",
  'Fashion & Luxury': "['Crafting Luxury Brand Narratives (Vogue Business 2023)']",
  'Technology': "['Tech Communication Trends in European Markets']",
  'Manufacturing': "['Industrial Communication Best Practices (Manufacturing Today 2024)']",
  'Finance': "['Financial Writing Standards in the Netherlands (European Finance Review)']",
  'Automotive': "['German Automotive Communication Standards']",
  'Design': "['Italian Design Language (Domus Magazine 2023)']",
  'Healthcare': "['Ethical Medical Communications in the Nordic Region']",
  'Tourism': "['Tourism Marketing Content Strategy for Southern Europe']",
  'Government': "['EU Public Sector Communication Guidelines']",
  'Education': "['Educational Content Localization in Central Europe']",
  'Retail': "['Nordic Retail Content Trends 2025']",
  'Style Guardian': "['The TextCraft Style Bible (Internal Publication)','Evolution of Editorial Standards in Digital Publishing']",
  'Fiction & Poetry': "['The Art of French Literary Editing (Le Monde des Livres)']",
  'Documentation': "['German Technical Documentation Standards (tekom Journal)']",
  'Advertising': "['Italian Advertising Copy: Art Meets Commerce']",
  'Research Papers': "['Nordic Academic Editing Standards (Scandinavian Journal of Publishing)']",
  'Business Reports': "['Dutch Business Report Writing Best Practices']",
  'Translation Quality': "['Translation Quality Metrics for European Markets']",
  'Localisation': "['Localisation Challenges in Celtic Languages']",
  'Legal & Formal': "['Austrian Legal Language Modernization']",
  'Global Standards': "['Swiss Multilingual Standards in Corporate Communication']",
  'Research & Journals': "['Academic Publishing Ethics (Oxford University Press 2024)','The Peer Review Process: A Critical Analysis']",
  'Scientific Papers': "['Chemical Nomenclature in Scientific Publishing']",
  'Medical Research': "['Medical Writing Ethics in Clinical Trials (Nordic Medical Journal)']",
  'Social Sciences': "['Qualitative Research Writing Methods (European Sociological Review)']",
  'Engineering': "['Technical Writing for Engineering Specifications']",
  'Environmental Studies': "['Environmental Impact Assessment Writing Standards']",
  'Economics': "['Economic Report Writing in the Eurozone']",
  'Literature Review': "['Systematic Literature Review Methodology']",
  'Annual Reports': "['Annual Report Design Trends in DACH Region']",
  'Sustainability Reports': "['GRI Sustainability Reporting in Practice']",
  'EU Communications': "['EU Regulation Communication for Citizens']",
  'Press Releases': "['Press Release Writing in the Digital Age']",
  'Internal Communications': "['Internal Communications During Remote Work']",
  'Investor Relations': "['Investor Relations Communication Standards']",
  'Executive Speeches': "['The Art of Executive Speechwriting']",
  'Banking & Finance': "['Swiss Banking Communication Compliance']",
  'Software Documentation': "['Modern Software Documentation Practices']",
  'API Documentation': "['RESTful API Documentation Standards']",
  'User Manuals': "['User-Centered Manual Design']",
  'Hardware Guides': "['EU Machinery Documentation Compliance']",
  'Engineering Specs': "['Engineering Specification Templates for EU Projects']",
  'Aerospace': "['Aerospace Technical Writing Standards (ESA Publications)']",
  'Telecom': "['5G Technology Documentation Standards']",
  'Fiction': "['The French Novel in the 21st Century (Les Inrockuptibles)','Voices of Modern Europe: An Anthology']",
  'Short Stories': "['The Dutch Short Story Renaissance']",
  'Poetry': "['Contemporary Irish Poetry and European Influences']",
  'Dramatic Writing': "['Italian Dramatic Writing: Stage to Screen']",
  'Nordic Noir': "['Nordic Noir Writing Workshop Handbook']",
  'Historical Fiction': "['Polish Historical Fiction: Memory and Identity']",
  'Contemporary Fiction': "['British Contemporary Fiction Trends']",
  'Magical Realism': "['Portuguese Magical Realism in European Context']",
  'Brand Campaigns': "['Brand Campaign Copywriting Handbook']",
  'Digital Marketing': "['German Digital Marketing Copy Optimization']",
  'Social Media': "['Social Media Copywriting for European Audiences']",
  'Video Scripts': "['Video Script Writing for Multilingual Markets']",
  'Retail Campaigns': "['Nordic Retail Campaign Case Studies']",
  'Luxury Brands': "['Luxury Brand Voice in Italian Markets']",
  'Tourism & Travel': "['Tourism Destination Marketing Content']",
  'Hospitality': "['Hospitality Industry Content Marketing']",
  'Session Management': "['Collaborative Review Process Design']",
  'Cross-Dept Sessions': "['Cross-Functional Team Facilitation Methods']",
  'Client Reviews': "['Client Feedback Integration in Creative Workflows']",
  'Creative Sessions': "['Creative Brainstorming Facilitation Techniques']",
  'Editorial Reviews': "['Editorial Review Process Optimization']",
  'Style Discussions': "['Multilingual Style Guide Development']",
  'Technical Reviews': "['Technical Review Automation in Publishing']",
  'Final Approvals': "['Quality Gate Process in Content Publishing']",
  'Standards': "['European Quality Standards in Publishing (QA Journal 2024)']",
  'French Texts': "['French Language Quality Assurance Methods']",
  'German Texts': "['German Text Proofreading Automation Study']",
  'Italian Texts': "['Italian Language Preservation in Corporate Texts']",
  'Spanish Texts': "['Spanish Language Variants in Business Communication']",
  'Polish Texts': "['Polish Language Modernization in Technical Texts']",
  'Research Verification': "['Fact-Checking Methodologies in the AI Era']",
  'Academic Verification': "['Academic Integrity Verification Tools Review']",
  'Style Compliance': "['Style Compliance Automation with NLP']",
  'Academic Integrity': "['Plagiarism Detection in Multilingual Texts']",
  'Operations': "['Operations Excellence in Creative Industries']",
  'People Operations': "['HR Best Practices for Creative Organizations']",
  'Budgets & Reporting': "['Financial Reporting for Creative Agencies']",
  'Infrastructure': "['IT Infrastructure for Distributed Creative Teams']",
  'User Support': "['IT Support Optimization for Creative Professionals']",
  'Recruitment': "['Recruiting Creative Talent in Europe']",
  'Payroll': "['Multi-Country Payroll Management in the EU']",
  'Facilities': "['Office Management for Creative Workspaces']",
  'Scheduling': "['Resource Scheduling for Multi-Office Organizations']",
  'Cloud Systems': "['Cloud Infrastructure for Publishing Workflows']",
};

// Patents - only relevant for tech, engineering, QA, and IT roles
const patentsData = {
  'Software Documentation': "['Automated Documentation Generation System (EP2024/001234)']",
  'API Documentation': "['Interactive API Documentation Framework (EP2023/005678)']",
  'Engineering Specs': "['Template-Based Engineering Specification Generator (EP2024/003456)']",
  'Aerospace': "['Aerospace Documentation Compliance Checker (EP2023/007890)']",
  'Infrastructure': "['Automated Publishing Pipeline Architecture (EP2024/002345)']",
  'Cloud Systems': "['Distributed Content Management System (EP2023/006789)']",
  'Standards': "['Multilingual Quality Assurance Scoring Engine (EP2024/004567)']",
  'Academic Integrity': "['Multi-Language Plagiarism Detection Algorithm (EP2023/008901)']",
  'Research Verification': "['Automated Fact Verification System (EP2024/001567)']",
  'Style Compliance': "['NLP-Based Style Compliance Analyzer (EP2023/009012)']",
  'Global Standards': "['Cross-Language Style Consistency Validator (EP2024/005890)']",
  'Digital Marketing': "['AI-Powered Copy Optimization Engine (EP2023/003456)']",
  'Technology': "['Real-Time Content Collaboration Platform (EP2024/007123)']",
  'User Support': "['Intelligent IT Ticket Routing System (EP2023/004567)']",
};

for (const row of records) {
  const spec = row.Specialization || '';
  const role = row.role || '';
  
  // Educational activities
  let edu = educationData[role];
  if (!edu) {
    edu = specEducation[spec] || "['Professional Development Programme']";
  }
  row.educationalActivities = edu;

  // Publications
  let pub = publicationsData[spec];
  if (!pub) {
    pub = "[]";
  }
  row.publications = pub;

  // Patents (most people won't have patents)
  let pat = patentsData[spec];
  if (!pat) {
    pat = "[]";
  }
  row.patents = pat;
}

// Get original headers and append new columns
const originalHeaders = Object.keys(records[0]).filter(h => 
  h !== 'educationalActivities' && h !== 'publications' && h !== 'patents'
);
const headers = [...originalHeaders, 'educationalActivities', 'publications', 'patents'];

// Simple CSV serializer that quotes fields containing commas, quotes, or brackets
function escapeField(val) {
  const s = String(val ?? '');
  if (s.includes(',') || s.includes('"') || s.includes('\n') || s.includes('[') || s.includes("'")) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

const lines = [headers.join(',')];
for (const row of records) {
  const fields = headers.map(h => escapeField(row[h]));
  lines.push(fields.join(','));
}

fs.writeFileSync(csvPath, lines.join('\n') + '\n');
console.log(`Updated ${records.length} rows with educationalActivities, publications, patents columns`);
