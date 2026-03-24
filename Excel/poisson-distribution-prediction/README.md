# Poisson Distribution Prediction in Excel

An Excel-based football match prediction model that uses the **Poisson Distribution** to estimate expected goals, match outcome probabilities, and implied betting odds from team performance data.

---

## Overview

This project showcases how Excel can be used for **predictive analytics** and **statistical modelling** in a practical sports context.

Using historical team performance data, I built a model that:

- calculates **attack** and **defense ratings**
- estimates **expected goals** for both teams
- applies the **Poisson Distribution** to predict scoring probabilities
- generates **Home Win**, **Away Win**, and **Draw** probabilities
- converts those probabilities into **betting-style decimal odds**

The model is designed to be simple, reusable, and interactive for football match prediction.

---

## Project Goal

The goal of this project was to create a structured Excel model that demonstrates my ability to combine:

- Excel modelling
- probability and statistics
- analytical thinking
- sports data analysis
- formula-based automation

This project reflects how spreadsheet tools can go beyond reporting and be used to build real analytical models.

---

## Workbook Structure

### 1. Data Source
This sheet contains the input data used in the model, including:

- League position
- Team name
- Games played
- Wins, draws, and losses
- Goals for
- Goals against
- Goal difference
- Points
- Attack rating
- Defense rating

These values form the foundation of the prediction logic.

### 2. Predictive Model
This sheet is the main analysis area where the model:

- allows a **home team** and **away team** to be selected
- retrieves team ratings dynamically
- calculates **expected goals**
- uses `POISSON.DIST` to estimate the probability of scoring 0, 1, 2, 3, and more goals
- aggregates score probabilities into:
  - Home Win probability
  - Away Win probability
  - Draw probability
- calculates implied decimal odds from those probabilities

---

## What I Built

In this Excel project, I:

- structured the source data into a usable analytical format
- created team **attack** and **defense** strength metrics
- built a match prediction model based on statistical probability
- used lookup formulas to dynamically return team values
- calculated expected goals for both teams
- modelled scoring likelihood using the **Poisson Distribution**
- converted probability outputs into match result predictions
- designed the workbook so it can be reused for different fixtures

---

## Excel Functions and Concepts Used

This project makes use of the following Excel concepts and functions:

- `XLOOKUP`
- `POISSON.DIST`
- probability modelling
- expected goals calculation
- attack and defense ratings
- dynamic team selection
- formula-driven prediction logic
- outcome probability aggregation

---

## Skills Demonstrated

This project highlights my skills in:

- Excel analytics
- predictive modelling
- statistical analysis
- spreadsheet design
- formula logic
- data structuring
- sports analytics
- problem solving

---

## Why This Project Matters

This project demonstrates that Excel can be used as more than a reporting tool. It can also support:

- analytical modelling
- forecasting
- probability-based decision making
- reusable business logic in spreadsheet form

It also shows my ability to take raw data, structure it clearly, and turn it into a practical model with meaningful outputs.

---

## File Included

- `poisson_dist_prediction_excel.xlsx` — main Excel project workbook

---

## Author

**Harone Javed**

This project is part of my portfolio and showcases my interest in **Excel analytics**, **statistical modelling**, and **predictive problem-solving**.
