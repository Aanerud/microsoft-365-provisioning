<#
.SYNOPSIS
    Updates textcraft-europe.csv to ensure all persons have values within specified ranges:
    educationalActivities (2-3), languages (1-3), interests (3-5), patents (0-3), publications (0-5)
#>

$csvPath = Join-Path $PSScriptRoot "..\config\textcraft-europe.csv"
$csv = Import-Csv $csvPath

# Helper to parse array-string like "['a','b']" into a list
function Parse-ArrayStr([string]$raw) {
    if (-not $raw -or $raw -eq '[]') { return @() }
    $raw = $raw.Trim()
    $raw = $raw -replace "^\[", ""
    $raw = $raw -replace "\]$", ""
    # Split on ',' while respecting content
    $items = @()
    $parts = $raw -split "','"
    foreach ($p in $parts) {
        $p = $p.Trim().Trim("'").Trim()
        if ($p) { $items += $p }
    }
    return $items
}

# Helper to format list back to "['a','b']"
function Format-ArrayStr([string[]]$items) {
    if (-not $items -or $items.Count -eq 0) { return "[]" }
    $joined = ($items | ForEach-Object { "'$_'" }) -join ","
    return "[$joined]"
}

# ═══════════════════════════════════════════════════════════════
# Define per-person changes keyed by employeeId
# ═══════════════════════════════════════════════════════════════

$changes = @{}

# --- LANGUAGES: Trim from >3 to 3 ---
$changes['TCE015'] = @{ # Pierre Lefevre - had 4
    languages_set = @("French (Native)","Dutch (Fluent)","English (Fluent)")
}
$changes['TCE024'] = @{ # Elena Rodriguez - had 5
    languages_set = @("Spanish (Native)","English (Fluent)","Portuguese (Professional)")
}
$changes['TCE038'] = @{ # Isabelle Mercier - had 4
    languages_set = @("French (Native)","Dutch (Fluent)","English (Fluent)")
}
$changes['TCE027'] = @{ # Lukas Fischer - had 4
    languages_set = @("German (Native)","English (Fluent)","French (Professional)")
}
$changes['TCE043'] = @{ # Karin Meier - had 4
    languages_set = @("German (Native)","English (Fluent)","French (Professional)")
}

# --- LANGUAGES: Add 2nd language for people with only 1 ---
$changes['TCE044'] = @{ # Oliver Hughes
    languages_set = @("English (Native)","French (Conversational)")
}
$changes['TCE058'] = @{ # George Crawford
    languages_set = @("English (Native)","Italian (Conversational)")
}
$changes['TCE088'] = @{ # Robert Taylor
    languages_set = @("English (Native)","German (Conversational)")
}
$changes['TCE095'] = @{ # Benjamin Moore
    languages_set = @("English (Native)","Irish (Conversational)")
}

# --- EDUCATIONALACTIVITIES: Fix Isabel Fernandez (had 1) ---
$changes['TCE093'] = @{
    edu_set = @("BA Business Administration - Universidad Complutense Madrid","IFMA Facility Management Professional Certificate")
}

# --- EDUCATIONALACTIVITIES: Add 3rd to ~30 people for variety ---
# Merge with existing changes where needed
function Merge-Change($id, $key, $val) {
    if (-not $script:changes.ContainsKey($id)) { $script:changes[$id] = @{} }
    $script:changes[$id][$key] = $val
}

# Executives + Directors get 3rd education
Merge-Change 'TCE001' 'edu_add' @("Stanford Graduate School of Business Executive Program")
Merge-Change 'TCE002' 'edu_add' @("PMP Certification - Project Management Institute")
Merge-Change 'TCE003' 'edu_add' @("INSEAD International Finance Program")
Merge-Change 'TCE004' 'edu_add' @("Berlin School of Creative Leadership Fellowship")
Merge-Change 'TCE005' 'edu_add' @("Chartered Institute of Marketing (CIM) Diploma")

# Key Account team leads/managers
Merge-Change 'TCE006' 'edu_add' @("Stanford Sales Leadership Program")
Merge-Change 'TCE008' 'edu_add' @("AWS Cloud Practitioner Certification")
Merge-Change 'TCE010' 'edu_add' @("CFA Institute Investment Foundations Program")

# Editorial leads
Merge-Change 'TCE018' 'edu_add' @("Cambridge Certificate in Advanced Publishing")
Merge-Change 'TCE019' 'edu_add' @("Paris Book Fair Professional Development Certificate")
Merge-Change 'TCE020' 'edu_add' @("DITA Specialist Certification - OASIS")

# Academic writing leads
Merge-Change 'TCE028' 'edu_add' @("Royal Literary Fund Fellowship")
Merge-Change 'TCE029' 'edu_add' @("Max Planck Society Science Writing Fellowship")
Merge-Change 'TCE030' 'edu_add' @("WHO Health Communication Certificate")

# Corporate writing leads
Merge-Change 'TCE036' 'edu_add' @("Frankfurt Book Fair Professional Communications Certificate")
Merge-Change 'TCE037' 'edu_add' @("GRI Standards Sustainability Certificate")

# Technical writing leads
Merge-Change 'TCE044' 'edu_add' @("MIT OpenCourseWare Technical Communication Certificate")
Merge-Change 'TCE045' 'edu_add' @("AWS Technical Documentation Certification")
Merge-Change 'TCE049' 'edu_add' @("BSc Aerospace Engineering - Toulouse")

# Literary writing leads
Merge-Change 'TCE052' 'edu_add' @("Iowa Writers Workshop Summer Fellowship")
Merge-Change 'TCE053' 'edu_add' @("European Short Story Prize Workshop")

# Marketing copy leads
Merge-Change 'TCE060' 'edu_add' @("Cannes Lions School Creative Leadership Certificate")
Merge-Change 'TCE061' 'edu_add' @("Facebook Blueprint Advanced Certification")

# Round Table + QA + Operations leads
Merge-Change 'TCE068' 'edu_add' @("London School of Economics Mediation Certificate")
Merge-Change 'TCE076' 'edu_add' @("BSc Publishing - London College of Communication")
Merge-Change 'TCE086' 'edu_add' @("Lean Management Institute Master Certificate")
Merge-Change 'TCE089' 'edu_add' @("CompTIA Security+ Certification")
Merge-Change 'TCE082' 'edu_add' @("BA Journalism - Universite Libre de Bruxelles")

# --- PATENTS: Add to ~25 more people for variety (currently ~12 have 1) ---
# Technical writers and engineers get more patents
Merge-Change 'TCE020' 'patents_set' @("Structured Authoring Optimization System (EP2024/006234)")
Merge-Change 'TCE022' 'patents_set' @("Academic Citation Auto-Formatter (EP2024/008123)")
Merge-Change 'TCE028' 'patents_set' @("Automated Academic Reference Validation System (EP2023/007234)")
Merge-Change 'TCE029' 'patents_set' @("Scientific Notation Standardization Tool (EP2024/009345)","STEM Content Accessibility Analyzer (EP2025/001234)")
Merge-Change 'TCE030' 'patents_set' @("Medical Writing Compliance Checker (EP2024/006789)")
Merge-Change 'TCE032' 'patents_set' @("Engineering Patent Documentation Formatter (EP2023/005567)")
Merge-Change 'TCE033' 'patents_set' @("Environmental Impact Text Analyzer (EP2024/003890)")
Merge-Change 'TCE036' 'patents_set' @("Annual Report Narrative Generator (EP2024/004123)")
Merge-Change 'TCE037' 'patents_set' @("ESG Data Visualization Narrative Engine (EP2025/000234)","Sustainability Metric Textual Summarizer (EP2024/008901)")
Merge-Change 'TCE044' 'patents_add' @("Developer Documentation Linting System (EP2025/000567)")
Merge-Change 'TCE045' 'patents_add' @("API Schema Documentation Auto-Generator (EP2024/007890)")
Merge-Change 'TCE046' 'patents_set' @("Context-Aware User Manual Generation System (EP2024/005678)")
Merge-Change 'TCE047' 'patents_set' @("Safety Document Multi-Language Validator (EP2025/001890)")
Merge-Change 'TCE048' 'patents_add' @("CAD-to-Documentation Bridge System (EP2025/002345)")
Merge-Change 'TCE049' 'patents_add' @("Aerospace Compliance Document Checker (EP2025/000123)")
Merge-Change 'TCE050' 'patents_set' @("Vehicle Systems Documentation Framework (EP2024/006543)")
Merge-Change 'TCE051' 'patents_set' @("Telecom Protocol Documentation Generator (EP2024/007654)")
Merge-Change 'TCE060' 'patents_set' @("Brand Voice Consistency Scoring Algorithm (EP2024/008765)")
Merge-Change 'TCE068' 'patents_set' @("Collaborative Review Session Workflow Engine (EP2024/005432)")
Merge-Change 'TCE076' 'patents_add' @("Content Quality Scoring Neural Network (EP2025/001678)")
Merge-Change 'TCE084' 'patents_add' @("Cross-Document Terminology Consistency Analyzer (EP2025/002567)")
Merge-Change 'TCE086' 'patents_set' @("Creative Operations Resource Allocation System (EP2024/003210)")
Merge-Change 'TCE089' 'patents_add' @("Distributed Publishing Pipeline Orchestrator (EP2025/003456)")
Merge-Change 'TCE090' 'patents_add' @("Self-Healing IT Infrastructure Monitor (EP2025/002890)")

# --- PUBLICATIONS: Enrich for variety (0-5) ---
# Academic writers get 3-5 publications
Merge-Change 'TCE028' 'pubs_set' @("Academic Publishing Ethics (Oxford University Press 2024)","The Peer Review Process: A Critical Analysis","Grant Writing Strategies for Humanities Researchers","Writing for Impact Journals: A Practical Guide")
Merge-Change 'TCE029' 'pubs_set' @("Chemical Nomenclature in Scientific Publishing","STEM Communication for Non-Expert Audiences","Clarity in Quantum Physics Writing (Nature Reviews 2024)")
Merge-Change 'TCE030' 'pubs_set' @("Medical Writing Ethics in Clinical Trials (Nordic Medical Journal)","Patient-Centered Medical Communication Standards","Nordic Healthcare Documentation Handbook")
Merge-Change 'TCE031' 'pubs_set' @("Qualitative Research Writing Methods (European Sociological Review)","Ethnographic Narrative Techniques in Academic Publishing","The Art of Social Science Dissemination")
Merge-Change 'TCE032' 'pubs_set' @("Technical Writing for Engineering Specifications","Patent Documentation Best Practices (IEEE Transactions 2024)","Women in Engineering Communication")
Merge-Change 'TCE033' 'pubs_set' @("Environmental Impact Assessment Writing Standards","Climate Communication for Policy Makers","Marine Science Narratives for Public Engagement")
Merge-Change 'TCE034' 'pubs_set' @("Economic Report Writing in the Eurozone","Austrian Economic Policy Communication Handbook","Data Visualization in Economic Reports")
Merge-Change 'TCE035' 'pubs_set' @("Systematic Literature Review Methodology","Meta-Analysis Reporting Standards Guide","Research Synthesis in the Digital Age")

# Literary writers get 2-4 publications
Merge-Change 'TCE052' 'pubs_set' @("The French Novel in the 21st Century (Les Inrockuptibles)","Voices of Modern Europe: An Anthology","Narrative Craft and Commercial Success","Fiction Workshop Methods for Professional Writers")
Merge-Change 'TCE053' 'pubs_set' @("The Dutch Short Story Renaissance","Flash Fiction: Form and Function","Experimental Narrative in Low Countries Literature")
Merge-Change 'TCE054' 'pubs_set' @("Contemporary Irish Poetry and European Influences","Lyrical Language in Brand Communication","The Spoken Word Movement in Ireland")
Merge-Change 'TCE055' 'pubs_set' @("Italian Dramatic Writing: Stage to Screen","Dialogue Techniques for Commercial Storytelling")
Merge-Change 'TCE056' 'pubs_set' @("Nordic Noir Writing Workshop Handbook","Atmosphere and Tension in Brand Storytelling","Danish Crime Fiction: A Cultural Analysis")
Merge-Change 'TCE057' 'pubs_set' @("Polish Historical Fiction: Memory and Identity","Archival Research Methods for Writers","Eastern European Narratives in Modern Publishing")
Merge-Change 'TCE058' 'pubs_set' @("British Contemporary Fiction Trends","Character-Driven Narrative in Commercial Writing","Social Realism in 21st Century British Fiction")
Merge-Change 'TCE059' 'pubs_set' @("Portuguese Magical Realism in European Context","Lusophone Literary Traditions in Translation","Storytelling and Saudade: A Writers Guide")

# Corporate writers get 2-3 publications
Merge-Change 'TCE036' 'pubs_set' @("Annual Report Design Trends in DACH Region","Executive Messaging in Times of Crisis","Corporate Storytelling for Investor Relations")
Merge-Change 'TCE037' 'pubs_set' @("GRI Sustainability Reporting in Practice","ESG Communication for Stakeholder Engagement","Dutch Sustainability Reporting Standards")
Merge-Change 'TCE038' 'pubs_set' @("EU Regulation Communication for Citizens","Multilingual Corporate Communication in Brussels","Policy Writing for Public Understanding")
Merge-Change 'TCE039' 'pubs_set' @("Press Release Writing in the Digital Age","Crisis Communication Case Studies in Italian Media")
Merge-Change 'TCE040' 'pubs_set' @("Internal Communications During Remote Work","Employee Engagement Through Storytelling","Nordic Corporate Culture Communication")
Merge-Change 'TCE041' 'pubs_set' @("Investor Relations Communication Standards","Financial Narrative Techniques for Earnings Reports")
Merge-Change 'TCE042' 'pubs_set' @("The Art of Executive Speechwriting","Rhetoric in Corporate Leadership Communication","Conference Presentation Writing Guide")
Merge-Change 'TCE043' 'pubs_set' @("Swiss Banking Communication Compliance","Private Banking Narrative Standards","Financial Writing Across Swiss Language Regions")

# Technical writers get 2-3 publications
Merge-Change 'TCE044' 'pubs_set' @("Modern Software Documentation Practices","Developer Experience Through Documentation","API Documentation That Developers Love (IEEE Software 2024)")
Merge-Change 'TCE045' 'pubs_set' @("RESTful API Documentation Standards","Developer Portal Design Patterns","OpenAPI Specification Best Practices")
Merge-Change 'TCE046' 'pubs_set' @("User-Centered Manual Design","Minimalist Documentation for Maximum Impact")
Merge-Change 'TCE047' 'pubs_set' @("EU Machinery Documentation Compliance","Safety Documentation Standards for European Markets")
Merge-Change 'TCE048' 'pubs_set' @("Engineering Specification Templates for EU Projects","CAD Documentation Integration Standards","Manufacturing Process Documentation Guide")
Merge-Change 'TCE049' 'pubs_set' @("Aerospace Technical Writing Standards (ESA Publications)","EASA Compliance Documentation Handbook","Safety-Critical Writing for Aviation")
Merge-Change 'TCE050' 'pubs_set' @("German Automotive Communication Standards","EV Technology Documentation Guide","Autonomous Vehicle Systems Writing")
Merge-Change 'TCE051' 'pubs_set' @("5G Technology Documentation Standards","Network Architecture Documentation Primer","Telecom Specification Writing Guide")

# Marketing copywriters get 1-3 publications
Merge-Change 'TCE060' 'pubs_set' @("Brand Campaign Copywriting Handbook","Creative Brief Writing for European Campaigns","Integrated Marketing Communication Guide")
Merge-Change 'TCE061' 'pubs_set' @("German Digital Marketing Copy Optimization","Conversion Copywriting: Science and Craft","Performance Marketing Language Patterns")
Merge-Change 'TCE062' 'pubs_set' @("Social Media Copywriting for European Audiences","Platform-Native Content Strategy Guide")
Merge-Change 'TCE063' 'pubs_set' @("Video Script Writing for Multilingual Markets","Commercial Storytelling Techniques")
Merge-Change 'TCE064' 'pubs_set' @("Nordic Retail Campaign Case Studies","Seasonal Marketing Copy Optimization")
Merge-Change 'TCE065' 'pubs_set' @("Luxury Brand Voice in Italian Markets","Aspirational Copywriting for Heritage Brands","The Language of European Luxury")
Merge-Change 'TCE066' 'pubs_set' @("Tourism Destination Marketing Content","Experience Marketing: Words That Inspire Travel")
Merge-Change 'TCE067' 'pubs_set' @("Hospitality Industry Content Marketing","Guest Experience Narratives for Hotels")

# Editors get 2-3 publications
Merge-Change 'TCE019' 'pubs_set' @("The Art of French Literary Editing (Le Monde des Livres)","Narrative Structure in Commercial Fiction","Poetry Editing: Preserving Voice While Polishing Craft")
Merge-Change 'TCE020' 'pubs_set' @("German Technical Documentation Standards (tekom Journal)","Information Architecture for Technical Content","Structured Authoring in the DACH Region")
Merge-Change 'TCE021' 'pubs_set' @("Italian Advertising Copy: Art Meets Commerce","Marketing Editorial Standards for European Brands")
Merge-Change 'TCE022' 'pubs_set' @("Nordic Academic Editing Standards (Scandinavian Journal of Publishing)","Research Paper Quality Metrics","Academic Writing Style Across Nordic Countries")
Merge-Change 'TCE023' 'pubs_set' @("Dutch Business Report Writing Best Practices","Corporate Editorial Standards for Annual Reports","Executive Communication Editing Guide")
Merge-Change 'TCE024' 'pubs_set' @("Translation Quality Metrics for European Markets","Multilingual Style Consistency Systems","Romance Language Localization Handbook")
Merge-Change 'TCE025' 'pubs_set' @("Localisation Challenges in Celtic Languages","English Variant Management for Global Brands","British vs American English: A Style Guide")
Merge-Change 'TCE026' 'pubs_set' @("Austrian Legal Language Modernization","DACH Region German Variants Guide","Legal German for Corporate Documents")
Merge-Change 'TCE027' 'pubs_set' @("Swiss Multilingual Standards in Corporate Communication","Neutral English for Global Audiences","Cross-Cultural Language Neutrality Guide")

# Key account managers get 1-2 publications
Merge-Change 'TCE007' 'pubs_set' @("Crafting Luxury Brand Narratives (Vogue Business 2023)","Fashion Content Strategy for European Houses")
Merge-Change 'TCE009' 'pubs_set' @("Industrial Communication Best Practices (Manufacturing Today 2024)","B2B Content Strategy for CEE Markets")
Merge-Change 'TCE010' 'pubs_set' @("Financial Writing Standards in the Netherlands (European Finance Review)","Compliance-Friendly Financial Content Guide")
Merge-Change 'TCE011' 'pubs_set' @("German Automotive Communication Standards","EV Transition Communication Best Practices")
Merge-Change 'TCE012' 'pubs_set' @("Italian Design Language (Domus Magazine 2023)","Design Communication for Global Markets")
Merge-Change 'TCE013' 'pubs_set' @("Ethical Medical Communications in the Nordic Region","Patient Communication Standards in Healthcare")
Merge-Change 'TCE014' 'pubs_set' @("Tourism Marketing Content Strategy for Southern Europe","Destination Storytelling: The Portuguese Approach")
Merge-Change 'TCE015' 'pubs_set' @("EU Public Sector Communication Guidelines","Multilingual Tender Documentation Standards")
Merge-Change 'TCE016' 'pubs_set' @("Educational Content Localization in Central Europe","E-Learning Content Quality Frameworks")
Merge-Change 'TCE017' 'pubs_set' @("Nordic Retail Content Trends 2025","Omnichannel Content Strategy for Retail Brands")

# Round table, QA, Operations get 1-3
Merge-Change 'TCE068' 'pubs_set' @("Collaborative Review Process Design","Facilitation Techniques for Creative Teams","Workshop Design for Content Organizations")
Merge-Change 'TCE069' 'pubs_set' @("Cross-Functional Team Facilitation Methods","Stakeholder Alignment in Creative Projects")
Merge-Change 'TCE070' 'pubs_set' @("Client Feedback Integration in Creative Workflows","Presentation Skills for Creative Reviews")
Merge-Change 'TCE071' 'pubs_set' @("Creative Brainstorming Facilitation Techniques","Design Thinking in Content Development")
Merge-Change 'TCE072' 'pubs_set' @("Editorial Review Process Optimization","Quality Assurance in Content Production")
Merge-Change 'TCE073' 'pubs_set' @("Multilingual Style Guide Development","Brand Voice Alignment Across Markets")
Merge-Change 'TCE074' 'pubs_set' @("Technical Review Automation in Publishing","Expert Verification Workflows for Accuracy")
Merge-Change 'TCE075' 'pubs_set' @("Quality Gate Process in Content Publishing","Final Review Best Practices for Agencies")
Merge-Change 'TCE076' 'pubs_set' @("European Quality Standards in Publishing (QA Journal 2024)","Quality Metrics for Creative Content","Building a Quality Culture in Publishing")
Merge-Change 'TCE077' 'pubs_set' @("French Language Quality Assurance Methods","Typographic Standards in Modern French Publishing")
Merge-Change 'TCE078' 'pubs_set' @("German Text Proofreading Automation Study","Duden Standards Compliance in Corporate Publishing")
Merge-Change 'TCE079' 'pubs_set' @("Italian Language Preservation in Corporate Texts","Regional Italian Variants in Business Communication")
Merge-Change 'TCE080' 'pubs_set' @("Spanish Language Variants in Business Communication","RAE Standards in Corporate Publishing")
Merge-Change 'TCE081' 'pubs_set' @("Polish Language Modernization in Technical Texts","Grammar Automation for Slavic Languages")
Merge-Change 'TCE082' 'pubs_set' @("Fact-Checking Methodologies in the AI Era","Source Verification Standards for Publishing","Automated Fact-Checking: Promises and Limits")
Merge-Change 'TCE083' 'pubs_set' @("Academic Integrity Verification Tools Review","Statistical Claims Verification Methods")
Merge-Change 'TCE084' 'pubs_set' @("Style Compliance Automation with NLP","Cross-Document Consistency Standards","Brand Voice Monitoring Systems")
Merge-Change 'TCE085' 'pubs_set' @("Plagiarism Detection in Multilingual Texts","Originality Verification in the Age of AI","Copyright Compliance for Content Agencies")
Merge-Change 'TCE086' 'pubs_set' @("Operations Excellence in Creative Industries","Scaling Creative Teams Across Europe")
Merge-Change 'TCE087' 'pubs_set' @("HR Best Practices for Creative Organizations","Talent Management in European Publishing")
Merge-Change 'TCE088' 'pubs_set' @("Financial Reporting for Creative Agencies","Project-Based Accounting in Publishing")
Merge-Change 'TCE089' 'pubs_set' @("IT Infrastructure for Distributed Creative Teams","Microsoft 365 Deployment for Publishing")
Merge-Change 'TCE090' 'pubs_set' @("IT Support Optimization for Creative Professionals","Intelligent Helpdesk Design for Knowledge Workers")
Merge-Change 'TCE091' 'pubs_set' @("Recruiting Creative Talent in Europe","Employer Branding for Publishing Companies")
Merge-Change 'TCE092' 'pubs_set' @("Multi-Country Payroll Management in the EU","Italian Tax Compliance for Creative Firms")
Merge-Change 'TCE093' 'pubs_set' @("Office Management for Creative Workspaces","Facility Design for Collaborative Teams")
Merge-Change 'TCE094' 'pubs_set' @("Resource Scheduling for Multi-Office Organizations","Administrative Efficiency in Publishing")
Merge-Change 'TCE095' 'pubs_set' @("Cloud Infrastructure for Publishing Workflows","Azure Administration for Content Platforms","Zero Trust Security for Creative Organizations")

# --- INTERESTS: Reduce some from 5 to 3-4 for variety ---
# ~20 people get trimmed to 3 or 4 interests
Merge-Change 'TCE009' 'interests_count' 4
Merge-Change 'TCE011' 'interests_count' 4
Merge-Change 'TCE013' 'interests_count' 4
Merge-Change 'TCE017' 'interests_count' 3
Merge-Change 'TCE023' 'interests_count' 4
Merge-Change 'TCE032' 'interests_count' 3
Merge-Change 'TCE034' 'interests_count' 4
Merge-Change 'TCE040' 'interests_count' 3
Merge-Change 'TCE046' 'interests_count' 4
Merge-Change 'TCE048' 'interests_count' 3
Merge-Change 'TCE050' 'interests_count' 4
Merge-Change 'TCE057' 'interests_count' 4
Merge-Change 'TCE062' 'interests_count' 4
Merge-Change 'TCE066' 'interests_count' 3
Merge-Change 'TCE069' 'interests_count' 4
Merge-Change 'TCE071' 'interests_count' 4
Merge-Change 'TCE074' 'interests_count' 3
Merge-Change 'TCE078' 'interests_count' 4
Merge-Change 'TCE081' 'interests_count' 4
Merge-Change 'TCE083' 'interests_count' 3
Merge-Change 'TCE091' 'interests_count' 4
Merge-Change 'TCE094' 'interests_count' 4

# ═══════════════════════════════════════════════════════════════
# Apply changes
# ═══════════════════════════════════════════════════════════════

foreach ($row in $csv) {
    $id = $row.employeeId
    if (-not $changes.ContainsKey($id)) { continue }
    $c = $changes[$id]

    # --- Languages ---
    if ($c.ContainsKey('languages_set')) {
        $row.languages = Format-ArrayStr $c['languages_set']
    }

    # --- educationalActivities ---
    if ($c.ContainsKey('edu_set')) {
        $row.educationalActivities = Format-ArrayStr $c['edu_set']
    }
    if ($c.ContainsKey('edu_add')) {
        $existing = Parse-ArrayStr $row.educationalActivities
        $new = $existing + $c['edu_add']
        $row.educationalActivities = Format-ArrayStr $new
    }

    # --- Patents ---
    if ($c.ContainsKey('patents_set')) {
        $row.patents = Format-ArrayStr $c['patents_set']
    }
    if ($c.ContainsKey('patents_add')) {
        $existing = Parse-ArrayStr $row.patents
        $new = $existing + $c['patents_add']
        $row.patents = Format-ArrayStr $new
    }

    # --- Publications ---
    if ($c.ContainsKey('pubs_set')) {
        $row.publications = Format-ArrayStr $c['pubs_set']
    }

    # --- Interests: trim to target count ---
    if ($c.ContainsKey('interests_count')) {
        $existing = Parse-ArrayStr $row.interests
        $target = $c['interests_count']
        if ($existing.Count -gt $target) {
            $trimmed = $existing[0..($target-1)]
            $row.interests = Format-ArrayStr $trimmed
        }
    }
}

# ═══════════════════════════════════════════════════════════════
# Write CSV back (preserving format)
# ═══════════════════════════════════════════════════════════════

$headers = ($csv[0].PSObject.Properties | ForEach-Object { $_.Name })
$lines = @()
$lines += ($headers -join ',')

foreach ($row in $csv) {
    $fields = @()
    foreach ($h in $headers) {
        $val = $row.$h
        if ($null -eq $val) { $val = '' }
        # Quote if contains comma, quote, or newline
        if ($val -match '[,"\r\n]') {
            $val = '"' + ($val -replace '"', '""') + '"'
        }
        $fields += $val
    }
    $lines += ($fields -join ',')
}

$output = $lines -join "`n"
[System.IO.File]::WriteAllText($csvPath, $output, [System.Text.UTF8Encoding]::new($false))

Write-Host "CSV updated successfully at: $csvPath" -ForegroundColor Green

# ═══════════════════════════════════════════════════════════════
# Validation
# ═══════════════════════════════════════════════════════════════

Write-Host "`nValidating..." -ForegroundColor Cyan
$verify = Import-Csv $csvPath
$errors = @()
foreach ($row in $verify) {
    $name = $row.name
    $edu = (Parse-ArrayStr $row.educationalActivities).Count
    $lang = (Parse-ArrayStr $row.languages).Count
    $int = (Parse-ArrayStr $row.interests).Count
    $pat = (Parse-ArrayStr $row.patents).Count
    $pub = (Parse-ArrayStr $row.publications).Count

    if ($edu -lt 2 -or $edu -gt 3) { $errors += "$name : educationalActivities=$edu (need 2-3)" }
    if ($lang -lt 1 -or $lang -gt 3) { $errors += "$name : languages=$lang (need 1-3)" }
    if ($int -lt 3 -or $int -gt 5) { $errors += "$name : interests=$int (need 3-5)" }
    if ($pat -lt 0 -or $pat -gt 3) { $errors += "$name : patents=$pat (need 0-3)" }
    if ($pub -lt 0 -or $pub -gt 5) { $errors += "$name : publications=$pub (need 0-5)" }
}

if ($errors.Count -eq 0) {
    Write-Host "All 95 people pass validation!" -ForegroundColor Green
} else {
    Write-Host "VALIDATION ERRORS:" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}

# Summary stats
Write-Host "`nDistribution:" -ForegroundColor Cyan
$verify | ForEach-Object {
    [PSCustomObject]@{
        Edu = (Parse-ArrayStr $_.educationalActivities).Count
        Lang = (Parse-ArrayStr $_.languages).Count
        Int = (Parse-ArrayStr $_.interests).Count
        Pat = (Parse-ArrayStr $_.patents).Count
        Pub = (Parse-ArrayStr $_.publications).Count
    }
} | ForEach-Object {
    $_.PSObject.Properties | ForEach-Object { [PSCustomObject]@{Field=$_.Name; Count=$_.Value} }
} | Group-Object Field | ForEach-Object {
    $field = $_.Name
    $counts = $_.Group | Group-Object Count | Sort-Object Name | ForEach-Object { "$($_.Name):$($_.Count)" }
    Write-Host "  $field  = $($counts -join ', ')"
}
