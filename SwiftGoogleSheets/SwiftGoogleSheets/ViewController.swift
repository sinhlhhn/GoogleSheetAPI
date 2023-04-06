//
//  ViewController.swift
//  SwiftGoogleSheets
//
//  Created by Roberto Sonzogni on 07/04/22.
//

import GoogleAPIClientForREST
import GoogleSignIn
import GTMSessionFetcher
import Toast_Swift
import UIKit

let spreadsheetId = Globals.shared.YOUR_SHEET_ID

class ViewController: UIViewController {
    override func viewDidLoad() {
        super.viewDidLoad()
        // Do any additional setup after loading the view.
    }
    
    @IBAction func signIn(sender: Any) {
        let signInConfig = GIDConfiguration(clientID: Globals.shared.YOUR_CLIENT_ID)
        GIDSignIn.sharedInstance.signIn(with: signInConfig, presenting: self) { user, error in
            guard error == nil else {
                self.view.makeToast("errore 1")
                return
            }
            
            // If sign in succeeded, display the app's main content View.
            guard
                let user = user
            else {
                self.view.makeToast("errore 2")
                return
            }
            
            // Your user is signed in!
            Globals.shared.gUser = user
            Globals.shared.sheetService.authorizer = Globals.shared.gUser.authentication.fetcherAuthorizer()
            print("logged in: \(user.profile?.name ?? "")")
            self.view.makeToast("logged in: \(user.profile?.name ?? "")")
        }
    }
    
    @IBAction func signOut(sender: Any) {
        GIDSignIn.sharedInstance.signOut()
        print("logged out")
        view.makeToast("logged out")
    }
    
    @IBAction func addScope() {
        let newScope = kGTLRAuthScopeSheetsSpreadsheets
        let grantedScopes = Globals.shared.gUser.grantedScopes
        
        if grantedScopes == nil || !grantedScopes!.contains(newScope) {
            view.makeToast("scope not present")
            
            // Request additional scope.
            let additionalScopes = [newScope]
            GIDSignIn.sharedInstance.addScopes(additionalScopes, presenting: self) { user, error in
                guard error == nil else {
                    self.view.makeToast("error 3")
                    return
                }
                guard let user = user else {
                    self.view.makeToast("error 4")
                    return
                }
                
                Globals.shared.gUser = user
                Globals.shared.sheetService.authorizer = Globals.shared.gUser.authentication.fetcherAuthorizer()
                
                // Check if the user granted access to the scopes you requested.
                let grantedScopes = Globals.shared.gUser.grantedScopes
                if grantedScopes!.contains(newScope) {
                    self.view.makeToast("scope added")
                }
            }
        } else {
            view.makeToast("scope already present")
        }
    }
    
    @IBAction func read() {
        let range = "B2"
        let query = GTLRSheetsQuery_SpreadsheetsValuesGet.query(withSpreadsheetId: Globals.shared.YOUR_SHEET_ID, range: range)
        Globals.shared.sheetService.executeQuery(query) { (_: GTLRServiceTicket, result: Any?, error: Error?) in
            if let error = error {
                print("Error", error.localizedDescription)
                self.view.makeToast(error.localizedDescription)
            } else {
                let data = result as? GTLRSheets_ValueRange
                let rows = data?.values as? [[String]] ?? [[""]]
                for row in rows {
                    print("row: ", row)
                    self.view.makeToast("value \(row.first ?? "") read in \(range)")
                }
                print("success")
            }
        }
    }
    
    @IBAction func write() {
        
        bathUpdate()
        
    }
    
    func bathUpdate() {
        let source = GTLRSheets_GridRange()
        source.sheetId = 0
        source.startRowIndex = 0
        source.endRowIndex = 7
        source.startColumnIndex = 0
        source.endColumnIndex = 1
        
        let sources = [source]
        
        let sourceRange = GTLRSheets_ChartSourceRange()
        sourceRange.sources = sources
        
        let chartData = GTLRSheets_ChartData()
        chartData.sourceRange = sourceRange
        
        let domain = GTLRSheets_BasicChartDomain()
        domain.domain = chartData
        
        let domains = [domain]
        
        let basicChart = GTLRSheets_BasicChartSpec()
        basicChart.domains = domains
        basicChart.chartType = "COLUMN"
        basicChart.legendPosition = "BOTTOM_LEGEND"
        
        let axis1 = GTLRSheets_BasicChartAxis()
        axis1.position = "BOTTOM_AXIS"
        axis1.title = "Categories"
        
        let axis2 = GTLRSheets_BasicChartAxis()
        axis2.position = "LEFT_AXIS"
        axis2.title = "Amount"
        
        let axiss = [axis1, axis2]
        basicChart.axis = axiss
        
        let series1 = GTLRSheets_BasicChartSeries()
        series1.targetAxis = "LEFT_AXIS"
        let chartData1 = GTLRSheets_ChartData()
        let chartSourceRange1 = GTLRSheets_ChartSourceRange()
        chartData1.sourceRange = chartSourceRange1
        let source1 = GTLRSheets_GridRange()
        source1.sheetId = 0
        source1.startRowIndex = 0
        source1.endRowIndex = 7
        source1.startColumnIndex = 1
        source1.endColumnIndex = 2
        chartSourceRange1.sources = [source1]
        series1.series = chartData1
        
        let series2 = GTLRSheets_BasicChartSeries()
        series2.targetAxis = "LEFT_AXIS"
        let chartData2 = GTLRSheets_ChartData()
        let chartSourceRange2 = GTLRSheets_ChartSourceRange()
        chartData2.sourceRange = chartSourceRange2
        let source2 = GTLRSheets_GridRange()
        source2.sheetId = 0
        source2.startRowIndex = 0
        source2.endRowIndex = 7
        source2.startColumnIndex = 2
        source2.endColumnIndex = 3
        chartSourceRange2.sources = [source2]
        series2.series = chartData2
        
        let series3 = GTLRSheets_BasicChartSeries()
        series3.targetAxis = "LEFT_AXIS"
        let chartData3 = GTLRSheets_ChartData()
        let chartSourceRange3 = GTLRSheets_ChartSourceRange()
        chartData3.sourceRange = chartSourceRange3
        let source3 = GTLRSheets_GridRange()
        source3.sheetId = 0
        source3.startRowIndex = 0
        source3.endRowIndex = 7
        source3.startColumnIndex = 3
        source3.endColumnIndex = 4
        chartSourceRange3.sources = [source3]
        series3.series = chartData3
        let series = [series1, series2, series3]
        
        basicChart.series = series
        basicChart.headerCount = 1
        
        let spec = GTLRSheets_ChartSpec()
        spec.basicChart = basicChart
        spec.title = "Sinhlh"
        
        let embededChart = GTLRSheets_EmbeddedChart()
        let objectPosition = GTLRSheets_EmbeddedObjectPosition()
        objectPosition.newSheet = 1
        embededChart.position = objectPosition
        embededChart.spec = spec
        
        let addChart = GTLRSheets_AddChartRequest()
        addChart.chart = embededChart
        
        let request = GTLRSheets_Request()
        request.addChart = addChart

        // create a batch update request
        let batchUpdate = GTLRSheets_BatchUpdateSpreadsheetRequest()
        batchUpdate.requests = [request]

        let query = GTLRSheetsQuery_SpreadsheetsBatchUpdate.query(withObject: batchUpdate, spreadsheetId: spreadsheetId)
        
        // execute the request
        Globals.shared.sheetService.executeQuery(query) { (ticket, response, error) in
          if let error = error {
            print("Error: \(error.localizedDescription)")
          } else {
            print("Chart added successfully.")
          }
        }
    }
    
    
    
    private func getSheetID() {
        
    }
    
    private func write1() {
        let range = "B2"
        let valueRange = GTLRSheets_ValueRange(json: [
            "majorDimension": "ROWS",
            "range": range,
            "values": [
                [
                    10,
                ],
            ],
        ])
        
        let query = GTLRSheetsQuery_SpreadsheetsValuesUpdate.query(withObject: valueRange, spreadsheetId: Globals.shared.YOUR_SHEET_ID, range: range)
        query.valueInputOption = "USER_ENTERED"
        query.includeValuesInResponse = true
        
        Globals.shared.sheetService.executeQuery(query) { (_: GTLRServiceTicket, result: Any?, error: Error?) in
            if let error = error {
                print("Error", error.localizedDescription)
                self.view.makeToast(error.localizedDescription)
            } else {
                let data = result as? GTLRSheets_UpdateValuesResponse
                let rows = data?.updatedData?.values as? [[String]] ?? [[""]]
                for row in rows {
                    print("row: ", row)
                    self.view.makeToast("value \(row.first ?? "") wrote in \(range)")
                }
                print("success")
            }
        }
    }
}
