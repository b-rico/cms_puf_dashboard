# ETL Script for Dashboard Data

This repository contains an ETL (Extract, Transform, Load) script used to process Excel data files and populate a SQL Server database. The resulting database is intended to be used as a data source for a Power BI dashboard.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Overview

This ETL script is designed to:
- Read Excel data files containing structured data.
- Transform and clean the data according to defined mappings and data type specifications.
- Create SQL Server database tables based on the configuration and mappings.
- Load the transformed data into the corresponding database tables.
- Add metadata columns for better data organization.

The resulting database is utilized as a data source for creating the cms_puf_dashboard using Power BI.

## Features

- Reads Excel files and processes data sheets.
- Transforms data values, including handling Yes/No values.
- Creates database tables based on configured mappings.
- Loads transformed data into database tables.
- Adds metadata columns for the measurement year and optional abbreviation.
- Checks for table existence before executing CREATE TABLE statement.

## Prerequisites

- Python 3.x
- Required Python packages: pandas, pyodbc, openpyxl

## Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/your-username/your-repo.git
   cd your-repo
