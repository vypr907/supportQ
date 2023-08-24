Attribute VB_Name = "variables"
'Just the Global variables

Global wb As Workbook
Global qSht As Worksheet
Global logSht As Worksheet
Global dataSht As Worksheet
Global searchSht As Worksheet

Public password As Variant
Public refID As Integer

Global startScreen As startScreenFrm
Global signIn As signInFrm
Global queueScreen As queueView
Global addUsrScreen As addUserFrm
Global authFrm As authFrm
Global reportView As reportFrm

Global lastUserRow As Integer
Global lastQRow As Integer
Global lastLogRow As Integer
Global usersRng As Range
Global authorized As Boolean
Global good2Go As Boolean

'variables for logSearch
Global tktState As Integer
Global startDate As String
Global endDate As String

'variables for manipulating temp xl instance and wb
Global temp As Workbook
Global xlApp As Excel.Application
