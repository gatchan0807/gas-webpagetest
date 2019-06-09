import WebPagetest = require('./WebPagetest')
import Utils = require('./Utils')

export const getTestResults = () => {
  const sheetNames = Utils.parseArrayValue(process.env.SHEET_NAME)
  const enabledNetworkErrorReport = Utils.parseBooleanNumberValue(process.env.NETWORK_ERROR_REPORT)
  if (!sheetNames) {
    throw new Error('should define SHEET_NAME in .env')
  }
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  if (!activeSpreadsheet) {
    throw new Error('Not found active spreadsheet')
  }
  sheetNames.forEach(sheetName => {
    const sheet = activeSpreadsheet.getSheetByName(sheetName)
    if (!sheet) {
      throw new Error(`Not found sheet by name:${sheetName}`)
    }
    const lastTestIdRow = Utils.getLastRow(sheet, 'A')
    const lastCompletedRow = Utils.getLastRow(sheet, 'B')
    Logger.log('lastTestIdRow: %s, lastCompletedRow: %s', lastTestIdRow, lastCompletedRow)
    if (lastTestIdRow === lastCompletedRow) {
      Logger.log('すべての testId の結果が取得済みです')
      return
    }
    const testIds = sheet
      .getRange(`A${lastCompletedRow + 1}:A${lastTestIdRow}`)
      .getValues()
      .reduce((a, b) => a.concat(b), [])

    if (!testIds.length) {
      Logger.log('対象 testId はありませんでした')
      return
    }
    Logger.log('testIds: %s', testIds.join('\n'))
    const wpt = new WebPagetest()
    const results = testIds.map(testId => {
      const results = wpt.results(testId)
      if (results instanceof Error) {
        Logger.log('Failed to fetch test result', results)
        if (enabledNetworkErrorReport) {
          throw results
        }
        // Just return empty results if ignore network error
        return wpt.createEmptyTestResults()
      }
      return results
    })

    const targetRange = sheet.getRange(lastCompletedRow + 1, 2, results.length, results[0].length)
    targetRange.setValues(results)
  })
}
