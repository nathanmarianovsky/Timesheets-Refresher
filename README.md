<h1 align="center">JS Implementation of Root-Finding Algorithms</h1>


# Table of Contents

- [Introduction](#introduction)
- [Setting Up](#setting-up)
- [Running the App](#running-the-app)


# Introduction

This simple app is designed to refresh the excel file hosting bi-weekly timesheets for company workers. The excel file consists of a tab per worker timesheet, each containing a start date, end date, and name associated to the worker along with the work hours. A typical layout expected for the app to function can be viewed in the sample excel file provided.  

# Setting Up
Inside the root directory of the project run:
```js
npm install
```
This will handle the installation of all node_modules.


# Running the App

Inside the root directory of the project run:
```js
node app.js
```
in order for the current timesheets to be cleared and updated to the next pay period as well as save the old timesheets to the directory which corresponds to the year of the current pay period.