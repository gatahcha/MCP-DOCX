[docxMCP-guide.md]
# LLM Guide: Creating High-Quality DOCX Documents with Credible Sources

## Overview
This guide enables LLMs to create professional, well-sourced DOCX documents using the FastMCP DOCX Document Creator server. Every document must include credible, clickable sources and deliver critical value through concise, accurate content.

## Core Principles

### 1. Source Credibility Requirements
**MANDATORY**: Every claim, statistic, or fact MUST be supported by a credible source.

**Acceptable Source Types (in order of preference):**
- Government agencies (.gov domains)
- Academic institutions (.edu domains)  
- Peer-reviewed journals and research publications
- Established news organizations (Reuters, AP, BBC, etc.)
- Industry reports from recognized organizations
- Official company reports and statements

**Unacceptable Sources:**
- Social media posts
- Personal blogs without credentials
- Wiki sites (except as starting points)
- Forums or discussion boards
- Sites with obvious bias or agenda
- Sources older than 2 years (unless historical context)

### 2. Content Value Standards
**Critical Value Delivery:**
- Address the "so what?" question for every section
- Provide actionable insights or clear implications
- Include specific data points, not generalizations
- Connect information to broader trends or impacts
- Offer forward-looking analysis when appropriate

**Conciseness Requirements:**
- Write just enough to convey the information clearly and completely
- Lead with the most important information
- Eliminate redundant or filler content
- Use specific numbers and dates
- Focus on unique insights, not common knowledge
- Each paragraph should serve a distinct purpose

## MCP Server Functionality Reference

### Core Document Management Tools

#### `create_new_document()`
**Purpose**: Initializes a new DOCX document with professional formatting
**Returns**: Confirmation message
**Technical Details**: 
- Sets Times New Roman as default font
- Applies standard margins (1" top/bottom, 0.75" left/right)
- Prepares document for content addition
**Usage**: Always call this first before any content addition

#### `get_document_status()`
**Purpose**: Check if document is active and view current title
**Returns**: Document status and title information
**When to Use**: Before adding content to verify document state

#### `save_document(filename)`
**Purpose**: Save the completed document as a DOCX file
**Parameters**: `filename` (string) - automatically adds .docx extension if missing
**Returns**: Success confirmation with file size
**Technical Details**: Creates a downloadable DOCX file with all formatting preserved

### Content Addition Tools

#### `add_document_title(title)`
**Purpose**: Add the main document title with premium formatting
**Parameters**: `title` (string)
**Formatting Applied**:
- 22pt Times New Roman font
- Bold styling
- Center alignment
- Full-width underline
**Best Practice**: Use specific, descriptive titles with scope/timeframe

#### `add_section_heading(heading)`
**Purpose**: Add section headers to organize content
**Parameters**: `heading` (string)
**Formatting Applied**:
- 18pt Times New Roman font
- NOT bold (clean professional look)
- Justified alignment
**Usage Strategy**: Create clear content hierarchy with descriptive headings

#### `add_paragraph_text(text)`
**Purpose**: Add body content paragraphs
**Parameters**: `text` (string)
**Formatting Applied**:
- 12pt Times New Roman font
- Justified alignment
- Professional paragraph spacing
**Content Guidelines**: Write substantive paragraphs that deliver complete thoughts and value

#### `add_source_citation(dates, full_source, shorten_source)`
**Purpose**: Add credible source citations with clickable links
**Parameters**:
- `dates` (string): Publication or access date
- `full_source` (string): Complete URL
- `shorten_source` (string): Display name (usually domain)
**Formatting Applied**:
- 10pt Times New Roman font
- Right-aligned positioning
- Clickable hyperlink functionality
- Format: (date, display_name)
**Critical Feature**: Creates actual clickable links in the DOCX file

#### `add_document_footer()`
**Purpose**: Add professional footer with document metadata
**Content**: Document title + copyright notice
**Formatting**: Centered alignment
**When to Use**: Call once at document completion, before saving

### Advanced Formatting Tools

#### `set_custom_margins(top, bottom, left, right)`
**Purpose**: Customize document margins beyond defaults
**Parameters**: All in inches (float values)
**Defaults**: top=1.0, bottom=1.0, left=0.75, right=0.75
**Use Cases**: Special formatting requirements or specific publication standards

#### `create_housing_report_example()`
**Purpose**: Generate a complete example document for reference
**Returns**: Fully formatted sample document
**Contains**: Title, sources, headings, paragraphs, and footer
**Usage**: Study the example structure for formatting guidance

### Tool Sequence Requirements
**Critical**: All content tools require an active document created with `create_new_document()`
**Error Handling**: Tools return specific error messages if no document exists
**Status Check**: Use `get_document_status()` to verify document state

### Understanding Tool Responses
**Success Indicators**: All successful operations return messages starting with "✅"
**Error Indicators**: Failed operations return messages starting with "❌"
**Information Tracking**: Tools provide feedback about what was added (e.g., font size, formatting)

**Example Successful Responses:**
- `"✅ New document created with Times New Roman formatting and standard margins"`
- `"✅ Title added: 'Vancouver Housing Report' (22pt, bold, centered with full-width underline)"`
- `"✅ Source added: (May 2025, cmhc-schl.gc.ca) - clickable link to https://..."`

**Example Error Response:**
- `"❌ No document created. Use create_new_document() first."`

**Response Analysis**: Pay attention to tool feedback to ensure proper formatting application and troubleshoot issues.

## Document Creation Workflow

### Phase 1: Research and Planning
Before using any MCP tools, complete this preparation:

1. **Topic Analysis**
   - Identify 3-5 key questions the document must answer
   - Determine the target audience and their needs
   - Set specific, measurable objectives for the document

2. **Source Discovery**
   - Compile 5-10 high-quality sources minimum
   - Verify publication dates (prefer recent sources)
   - Check author credentials and institutional affiliations
   - Ensure geographic/demographic relevance if applicable

3. **Content Architecture**
   - Create logical section flow
   - Prioritize information by importance
   - Plan 1-2 sources per major section

### Phase 2: Document Construction

#### Step 1: Initialize Document
```
create_new_document()
```

#### Step 2: Add Title
**Format**: Clear, specific, includes timeframe/scope
**Example**: "Vancouver Housing Affordability Crisis: Policy Solutions and Market Trends (Q2 2025)"
```
add_document_title("Your Specific Title Here")
```

#### Step 3: Build Content Sections

**For Each Section:**

1. **Add Section Heading** (specific, not generic)
   ```
   add_section_heading("Concrete Heading That Promises Value")
   ```

2. **Add Supporting Source**
   ```
   add_source_citation("Date", "Full URL", "Domain/Organization")
   ```

3. **Add Content Paragraph** (follow value standards)
   ```
   add_paragraph_text("Value-driven content...")
   ```

#### Step 4: Finalize Document
```
add_document_footer()
save_document("descriptive_filename")
```

## Content Writing Standards

### Opening Paragraphs
- Start with the most newsworthy or impactful information
- Include specific metrics or findings in the first sentence
- Answer "why does this matter?" within the first paragraph

### Section Structure
**Lead with Impact**: "Housing prices in Vancouver decreased 8.3% in Q1 2025, marking the steepest quarterly decline since 2008."

**Provide Context**: "This decline follows the implementation of the provincial Empty Homes Tax, which added a 3% annual levy on vacant properties. The policy targets speculative investment by imposing financial penalties on properties left vacant for more than six months annually."

**Deliver Insight**: "The trend suggests policy interventions can effectively cool speculative investment, potentially improving affordability for local buyers. Market analysts predict continued price corrections through 2025 as similar policies expand provincially."

**Paragraph Guidelines**:
- Write enough to fully develop the idea
- Avoid unnecessary elaboration or redundancy
- Each paragraph should advance the document's purpose
- Include supporting details that add genuine value

### Data Presentation
- Use specific numbers: "47% increase" not "significant increase"
- Include timeframes: "since March 2024" not "recently"
- Provide comparisons: "compared to the national average of 12%"
- Show trends: "continuing a three-month decline"

### Source Integration Best Practices

#### Source Citation Format
```
add_source_citation(
    dates="March 15, 2025", 
    full_source="https://www.cmhc-schl.gc.ca/housing-market-data/vancouver-q1-2025",
    shorten_source="cmhc-schl.gc.ca"
)
```

#### Source Selection Criteria
- **Timeliness**: Published within 6 months for current events
- **Authority**: Author has relevant expertise or institutional backing
- **Transparency**: Methodology clearly explained
- **Independence**: Free from obvious conflicts of interest

## Quality Assurance Checklist

### Before Document Creation:
- [ ] All sources verified for credibility and current relevance
- [ ] Content plan addresses specific audience needs
- [ ] Each section delivers unique value
- [ ] Statistics and claims are current and accurate
- [ ] Document purpose and scope clearly defined

### During Content Creation:
- [ ] Called `create_new_document()` before adding any content
- [ ] Every factual claim has a supporting source with `add_source_citation()`
- [ ] Paragraphs are substantive but eliminate unnecessary words
- [ ] Content answers "so what?" question
- [ ] Technical terms are explained appropriately
- [ ] Proper sequence: title → content sections → footer → save

### MCP Tool Usage Verification:
- [ ] Document title uses `add_document_title()` with descriptive text
- [ ] Section headings use `add_section_heading()` for clear organization
- [ ] All sources use `add_source_citation()` with complete URLs
- [ ] Body content uses `add_paragraph_text()` for substantial information
- [ ] Footer added with `add_document_footer()` before saving
- [ ] Document saved with `save_document()` using descriptive filename

### After Document Completion:
- [ ] Title accurately reflects content scope and timeframe
- [ ] Sources are properly formatted and create clickable links
- [ ] Document provides actionable insights
- [ ] Content is free of redundancy and filler
- [ ] Each section contributes meaningfully to the overall purpose

## Example Templates

### News Analysis Document
```
Title: "Federal Housing Policy Changes Impact Vancouver Market (June 2025)"
Source: Government of Canada housing announcements
Content Focus: Policy implications, market predictions, stakeholder impacts
```

### Research Report
```
Title: "Tech Sector Recovery Patterns: Vancouver vs. Toronto Analysis (Q2 2025)"
Sources: Statistics Canada, company earnings reports, industry surveys
Content Focus: Comparative analysis, trend identification, future outlook
```

### Policy Brief
```
Title: "Municipal Zoning Reform Effectiveness: Six-Month Assessment (2025)"
Sources: City planning documents, housing starts data, expert interviews
Content Focus: Implementation results, challenges, recommendations
```

## Accuracy and Conciseness Standards

### Preventing Information Bloat
**Content Filtering Principles:**
- Include only information that directly serves the document's purpose
- Eliminate background information readers likely already know
- Avoid repetitive explanations across sections
- Cut transitional phrases that don't add meaning
- Remove qualifiers that weaken statements without adding accuracy

**Value-Per-Word Optimization:**
- Every sentence should either inform, analyze, or provide context
- Combine related points rather than spreading across multiple paragraphs
- Use specific examples instead of abstract explanations
- Replace general statements with concrete data points

### Accuracy Verification Process
**Source Validation:**
- Cross-reference claims with primary sources
- Verify publication dates and author credentials
- Check for updated information or corrections
- Ensure statistical context (sample sizes, methodologies)

**Fact-Checking Protocol:**
- Distinguish between correlation and causation
- Include confidence levels for predictions
- Specify geographic and temporal scope of claims
- Note limitations or uncertainties in data

**MCP Tool Accuracy:**
- Use exact URLs in `add_source_citation()` - no shortened or redirect links
- Include complete dates (month/day/year when available)
- Use precise, descriptive text for headings and titles
- Ensure source display names match the actual organization

## Common Pitfalls to Avoid

### Content Quality Issues
1. **Generic Sourcing**: Don't cite sources just to meet requirements
2. **Information Dumping**: Avoid listing facts without analysis or connection
3. **Outdated Data**: Don't use sources older than context requires
4. **Weak Headlines**: Avoid vague section titles like "Background Information"
5. **Redundant Content**: Don't repeat information across sections
6. **Missing Context**: Don't present data without explaining significance
7. **Filler Content**: Eliminate sentences that don't advance understanding

### MCP Tool Usage Errors
1. **Skipping Document Creation**: Always call `create_new_document()` first
2. **Incorrect Source Formatting**: Use complete URLs and accurate dates in citations
3. **Poor File Naming**: Use descriptive filenames that reflect content and date
4. **Missing Footer**: Always add footer before saving for professional appearance
5. **Ignoring Tool Feedback**: Check success/error messages to ensure proper execution
6. **Wrong Tool Selection**: Use `add_section_heading()` for headers, not `add_paragraph_text()`

## Success Metrics

**A successful document:**
- Answers specific questions with credible evidence
- Provides insights not easily found elsewhere
- Includes actionable information or clear implications
- Uses current, authoritative sources throughout
- Delivers value proportional to reader time investment

Remember: Every sentence should either provide new information, add context, or deliver insight. If it doesn't meet this standard, delete it.