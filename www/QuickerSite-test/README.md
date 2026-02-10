<h1>QuickerSite - CMS for Windows Servers</h1>

## Introduction

QuickerSite (<a href="https://quickersite.com/" target="_blank">https://quickersite.com/</a>) is an all-in-one CMS (Content Management System) built in ASP/VBScript. QuickerSite runs on both Windows Servers (IIS 5 and higher) and most other Windows OS versions through IIS Express (Visual Studio Code, Visual Studio, Webmatrix, etc). QuickerSite also requires ASP.NET (2.0) installed on your host. 

## Demo

The QuickerSite demosite showcases most features: <a href="https://demo.quickersite.com/" target="_blank">https://demo.quickersite.com/</a>

## Installation
Install this application in the root of an IIS website and make sure IUSR has full permissions to the entire folder.
Make sure "default.asp" is the default document. The default installation uses an Access database and therefore 32bits-support needs to be enabled for your IIS application pool.
When browsing to your site the first time, you will be prompted for a password. By default this is "admin". You have to change it to something else. 

## Templating engine

You can use <a href="https://jstemplates.com/" target="_blank">jstemplates.com</a> to generate responsive web designs to use in QuickerSite. 
