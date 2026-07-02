# BuildClinical → Participant Tracker

A browser-based automation tool that migrates participant records from BuildClinical CSV exports into a structured Excel tracker. Built for the Courtney Lab at UC San Diego Health.

## The Problem

Our clinical research team was spending 5-6 hours every week manually copying participant data from BuildClinical into our Excel recruitment tracker. I built this to fix that.

## What It Does

Upload your BuildClinical CSV export and Excel tracker, hit Run, and the tool automatically routes each participant to the correct sheet based on their enrollment status, checks for duplicate phone numbers, and formats everything with color-coded rows.

Eligible participants go to Potential Participants, enrolled participants go to Scheduled and Active, and ineligible participants go to Ineligible — all in seconds.

## Results

Since deploying this in March 2026 the team went from 5-6 hours of weekly data entry down to under one hour. Participant recruitment went from ~20 per quarter to ~50.

## Stack

Python · Flask · openpyxl · pandas

## Setup
pip install flask openpyxl pandas

Then create this folder structure:
buildclinical-to-tracker/
app.py
templates/
index.html
uploads/
outputs/

Run:
python3 app.py

Open your browser and go to http://localhost:5050

## Note

Sample data files are not included in this repo to protect participant privacy. Available upon request.

## Built By

Samin Chowdhury — Courtney Lab, UC San Diego Health
