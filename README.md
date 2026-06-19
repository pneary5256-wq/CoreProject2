# CoreProject2

CoreProject2 is a lightweight web application scaffold for processing CSV exports from Aptem and determining applicant eligibility against configurable criteria.

## Purpose

The application will help teams quickly assess whether applicants meet baseline requirements (for example minimum age), based on information submitted in Aptem-generated CSV files.

## Planned Features

- Upload and read Aptem CSV files in the browser
- Parse and normalize applicant background data
- Apply configurable eligibility criteria
- Display clear pass/fail outcomes with supporting details
- Keep criteria configuration easy to update

## Development Roadmap

1. **Project scaffolding** (current stage)
   - Establish folder structure and placeholder modules
   - Set up static HTML/CSS/JavaScript entry points
2. **CSV ingestion and parsing**
   - Add robust CSV file loading and validation
   - Map CSV columns to internal data structures
3. **Eligibility engine**
   - Introduce configurable rule definitions
   - Evaluate applicant data against rules
4. **Results and usability**
   - Improve interface for uploads and results review
   - Add error handling and user feedback
5. **Quality and hardening**
   - Expand tests and validation scenarios
   - Strengthen edge-case handling and documentation
