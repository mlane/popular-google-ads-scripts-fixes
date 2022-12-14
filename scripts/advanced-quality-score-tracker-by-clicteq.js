/**
 *
 * Advanced Quality Score Tracker by Clicteq
 * The scripts plots Quality Score and impression weighed Quality Score on a daily basis
 * It has 3 graphs to show the split by expected CTR, landing page experience and ad relevance
 *
 * Version: 1.3
 * maintained by Clicteq
 *
 **/

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~//
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~//

//Control Panel

//ID of Google Sheet. Create a new Google sheet that you want the data to be outputted to and paste the URL between then quotation marks below
// UPDATE this to the actual Google Doc ID
// SEE: https://stackoverflow.com/a/35210316
const ID = 'DOCUMENT_ID'

const LIMIT = false
// const LIMIT = 5

const QUERY_START =
  'SELECT metrics.impressions, ad_group_criterion.quality_info.quality_score FROM keyword_view'

const QUERY_END = `segments.date DURING YESTERDAY${
  LIMIT ? ' LIMIT ' + LIMIT : ''
}`

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~//
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~//

function getDetailedReport(enums, metrics) {
  //Downloads a report with QS components and calculates weighted averages. Returns a sheet-friendly row.

  var result = {}
  result = {
    SearchPredictedCtr_BELOW_AVERAGE: 1,
    SearchPredictedCtr_AVERAGE: 2,
    0: 3,
  }
  var total = 0
  var total_impr = 0
  for (metric in metrics) {
    for (item in enums) {
      var m = metrics[metric]
      var e = enums[item]
      console.log('metric', m)
      console.log('enums', e)
      var query = `${QUERY_START} WHERE ${m} = ${e} AND ${QUERY_END}`
      var temp = query.replace('%m', m).replace('%e', e)
      var report = AdWordsApp.report(temp).rows()
      var i = 0
      var impr = 0

      while (report.hasNext()) {
        row = report.next()
        i = i + 1
        var impr_part = parseInt(row['metrics.impressions'])
        impr = impr + impr_part
      }
      result[metrics[metric] + '_' + enums[item]] = {
        impressions: parseInt(impr),
        total: parseInt(i),
      }
      total = total + i
      total_impr = total_impr + impr
    }
  }
  total = parseInt(total / 3)
  total_impr = parseInt(total_impr / 3)

  var sheetRow = []

  for (metric in metrics) {
    for (item in enums) {
      m = metrics[metric]
      e = enums[item]
      var single = result[m + '_' + e]
      single['total'] = (single['total'] / total).toFixed(5)
      sheetRow.push(result[m + '_' + e]['total'])
      if (e == 'BELOW_AVERAGE') {
        var weighted = 1 * (single['metrics.impressions'] / total_impr)
      } else if (e == 'AVERAGE') {
        var weighted = 2 * (single['metrics.impressions'] / total_impr)
      } else if (e == 'ABOVE_AVERAGE') {
        var weighted = 3 * (single['metrics.impressions'] / total_impr)
      }

      single['weighted'] = weighted
    }
  }

  var weightedAll = {}

  for (metric in metrics) {
    m = metrics[metric]
    weightedAll[m] = 0
    for (item in enums) {
      e = enums[item]
      var single = result[m + '_' + e]
      weightedAll[m] = weightedAll[m] + single['weighted']
    }
    weightedAll[m] = weightedAll[m].toFixed(4)
  }
  var metrics_temp = [
    'ad_group_criterion.quality_info.search_predicted_ctr',
    'ad_group_criterion.quality_info.post_click_quality_score',
    'ad_group_criterion.quality_info.creative_quality_score',
  ]
  for (item in metrics_temp) {
    sheetRow.push(weightedAll[metrics_temp[item]])
  }
  return sheetRow
}

function getTotalReport() {
  //Downloads a QS report and calculates weighted average. Returns a sheet-friendly row.
  var query = `${QUERY_START} WHERE ${QUERY_END}`
  var report = AdWordsApp.report(query).rows()
  var result = {
    1: { impressions: 0, total: 0, weighted: 0 },
    2: { impressions: 0, total: 0, weighted: 0 },
    3: { impressions: 0, total: 0, weighted: 0 },
    4: { impressions: 0, total: 0, weighted: 0 },
    5: { impressions: 0, total: 0, weighted: 0 },
    6: { impressions: 0, total: 0, weighted: 0 },
    7: { impressions: 0, total: 0, weighted: 0 },
    8: { impressions: 0, total: 0, weighted: 0 },
    9: { impressions: 0, total: 0, weighted: 0 },
    10: { impressions: 0, total: 0, weighted: 0 },
  }
  var impr = 0
  var total = 0
  var totalQs = 0
  while (report.hasNext()) {
    var row = report.next()
    var impr_part = parseInt(row['metrics.impressions'])
    var qs = row['ad_group_criterion.quality_info.quality_score'] ?? 1
    result[qs]['metrics.impressions'] =
      result[qs]['metrics.impressions'] + impr_part
    result[qs]['total'] = result[qs]['total'] + 1
    impr = impr + impr_part
    var total = total + 1
    totalQs = totalQs + parseInt(qs)
  }
  var temp = []
  var keys = Object.keys(result)
  var weighted = 0
  for (place in keys) {
    var key = keys[place]
    var weighted_part = (result[key]['metrics.impressions'] / impr) * key
    result[key]['total'] = (result[key]['total'] / total).toFixed(4)
    weighted = weighted + weighted_part
  }
  weighted = weighted.toFixed(2)
  var qses = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10']
  for (place in qses) {
    var key = qses[place]
    temp.push(result[key]['total'])
  }

  temp.push(weighted)
  temp.push((totalQs / total).toFixed(2))

  return temp
}

function createGraphs() {
  //Creates graphs when the script first runs

  var graphsData = {
    ctr: { all: 'A2:C4', weighted: 'K:K' },
    lp: { all: 'A5:C7', weighted: 'L:L' },
    ad: { all: 'A8:C10', weighted: 'M:M' },
    all: { all: 'A14:C23', weighted: 'X:X' },
  }
  var titles = {
    ctr: {
      all: 'QS expected CTR % of all keywords',
      weighted: 'QS expected CTR weighted impressions',
    },
    lp: {
      all: 'QS Landing Page % of all keywords',
      weighted: 'QS Landing Page weighted impressions',
    },
    ad: {
      all: 'QS Ad Relevance % of all keywords',
      weighted: 'QS Ad Relevance weighted impressions',
    },
    all: { all: 'QS % of all keywords', weighted: 'QS weighted impressions' },
  }
  var axes = {
    all_all: [0, 1],
    ctr_all: [0, 1],
    lp_all: [0, 1],
    ad_all: [0, 1],
    all_weighted: [0, 10],
    ctr_weighted: [0, 3],
    ad_weighted: [0, 3],
    lp_weighted: [0, 3],
  }
  var positions = [1, 4, 21, 38]

  var sheet = SpreadsheetApp.open(DriveApp.getFileById(ID)).getSheetByName(
    'dashboard'
  )
  var data = SpreadsheetApp.open(DriveApp.getFileById(ID)).getSheetByName(
    'data'
  )
  var types = ['all', 'ctr', 'lp', 'ad']
  var subTypes = ['all', 'weighted']
  var xAxis = data.getRange('A1:A')
  var dataBar = SpreadsheetApp.open(DriveApp.getFileById(ID)).getSheetByName(
    'data-bar'
  )
  var xAxisBar = dataBar.getRange('A1:C1')

  if (sheet.getCharts().length == 0) {
    Logger.log('Creating graphs...')
  } else {
    return
  }
  for (type in types) {
    for (subType in subTypes) {
      var seriesBar = dataBar.getRange(
        graphsData[types[type]][subTypes[subType]]
      )
      var series = data.getRange(graphsData[types[type]][subTypes[subType]])
      if (subType != 0) {
        var pos = subType * 9
      } else {
        var pos = 2
      }

      if (subTypes[subType] == 'all') {
        var chart = sheet
          .newChart()
          .asColumnChart()
          .setOption('vAxes', {
            0: {
              viewWindow: {
                min: axes[types[type] + '_' + subTypes[subType]][0],
                max: axes[types[type] + '_' + subTypes[subType]][1],
              },
            },
          })
          .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
          .addRange(xAxisBar)
          .addRange(seriesBar)
          .setPosition(positions[type], pos, 2, 2)
          .setNumHeaders(1)
          .setOption('title', titles[types[type]][subTypes[subType]])
          .setOption('legend', { position: 'bottom' })
          .build()
      } else {
        var chart = sheet
          .newChart()
          .asLineChart()
          .addRange(xAxis)
          .addRange(series)
          .setPosition(positions[type], pos, 2, 2)
          .setNumHeaders(1)
          .setOption('title', titles[types[type]][subTypes[subType]])
          .setOption('legend', { position: 'bottom' })
          .setOption('vAxes', {
            0: {
              viewWindow: {
                min: axes[types[type] + '_' + subTypes[subType]][0],
                max: axes[types[type] + '_' + subTypes[subType]][1],
              },
            },
          })
          .build()
      }
      sheet.insertChart(chart)
    }
  }
}

function prepareSheet() {
  //Prepares a new sheet by creating new tabs

  var sheet = SpreadsheetApp.open(DriveApp.getFileById(ID))
  if (sheet.getSheetByName('dashboard') === null) {
    try {
      sheet.insertSheet('dashboard')
    } catch (e) {}
    try {
      sheet.insertSheet('data')
      sheet.getSheetByName('data').insertColumns(1, 12)
    } catch (e) {}
    try {
      sheet.insertSheet('data-bar')
      sheet.getSheetByName('data').insertColumnAfter(1)
    } catch (e) {}
    Logger.log('Preparing sheet...')
    try {
      sheet.deleteSheet(sheet.getSheetByName('Sheet1'))
    } catch (e) {}

    sheet.getSheetByName('dashboard').setHiddenGridlines(true)
    formatCells()
  }
}
function formatRange() {
  //Formats decimals so they are displayed as %

  var sheet = SpreadsheetApp.open(DriveApp.getFileById(ID))
  sheet.getRange('data!B2:J').setNumberFormat('0%')
  sheet.getRange('data!N2:W').setNumberFormat('0%')
  sheet.getRange('data!A2:A').setNumberFormat('dd/mm/yyyy')
}

function transpose() {
  //Creates a transposed version of the first and last days, used for bar charts

  var sheet = SpreadsheetApp.open(DriveApp.getFileById(ID)).getSheetByName(
    'data-bar'
  )
  var range = sheet.getRange('A:C')
  range.clear()
  var lastRow = SpreadsheetApp.open(DriveApp.getFileById(ID))
    .getSheetByName('data')
    .getLastRow()
  sheet.getRange('A1:A1').setFormula('=TRANSPOSE(data!A1:AD1)')
  sheet.getRange('B1:B1').setFormula('=TRANSPOSE(data!A2:AD2)')
  sheet
    .getRange('C1:C1')
    .setFormula('=TRANSPOSE(data!A' + lastRow + ':AD' + lastRow + ')')
}

function formatCells() {
  //Formats cells' size and fonts in the dashboard tab. Adds average quality scores as text

  var sheet = SpreadsheetApp.open(DriveApp.getFileById(ID)).getSheetByName(
    'dashboard'
  )
  var rows = [1]
  var cols_avg = ['Y', 'X']
  var cols = ['H', 'O']
  var avr = ['Normal average', 'Weighted average']
  for (row in rows) {
    sheet.setRowHeight(rows[row], 75)
    sheet.setRowHeight(rows[row] + 2, 280)
    for (col in cols) {
      var _col = cols[col]
      sheet
        .getRange(_col + rows[row] + ':' + _col + rows[row])
        .setValue('Average QS')
        .setFontWeight('bold')
        .setFontSize(24)
      sheet
        .getRange(_col + (rows[row] + 1) + ':' + _col + (rows[row] + 1))
        .setValue(avr[col])
      sheet
        .getRange(_col + (rows[row] + 2) + ':' + _col + (rows[row] + 2))
        .setFormula(
          '=AVERAGE(data!' + cols_avg[col] + '2:' + cols_avg[col] + ')'
        )
        .setNumberFormat('0.00')
        .setFontWeight('bold')
        .setFontSize(24)
      sheet
        .getRange(_col + rows[row] + ':' + _col + (rows[row] + 2))
        .setBorder(true, false, true, true, null, true)

      sheet
        .getRange(_col + ':' + _col)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
    }
  }
  sheet.autoResizeColumn(8)
  sheet.autoResizeColumn(15)
}

function main() {
  //Combines all functions
  prepareSheet()

  var enums = ['BELOW_AVERAGE', 'AVERAGE', 'ABOVE_AVERAGE']
  var metrics = [
    'ad_group_criterion.quality_info.search_predicted_ctr',
    'ad_group_criterion.quality_info.post_click_quality_score',
    'ad_group_criterion.quality_info.creative_quality_score',
  ]
  var sheet = SpreadsheetApp.open(DriveApp.getFileById(ID)).getSheetByName(
    'data'
  )

  Logger.log('Downloading detailed quality scores...')
  var detailedReport = getDetailedReport(enums, metrics)
  Logger.log('Downloading combined quality scores...')
  var totalReport = getTotalReport()

  var header = [
    [
      'Date',
      'Expected_CTR_Below_Average',
      'Expected_CTR_Average',
      'Expected_CTR_Above_Average',
      'LP_Below_Average',
      'LP_Average',
      'LP_Above_Average',
      'Ad_Copy_Below_Average',
      'Ad_Copy_Average',
      'Ad_Copy_Above_Average',
      'CTR_weighted',
      'LP_weighted',
      'Ad_Copy_weighted',
      '="1"',
      '="2"',
      '="3"',
      '="4"',
      '="5"',
      '="6"',
      '="7"',
      '="8"',
      '="9"',
      '="10"',
      'QS_weighted',
      'QS_average',
    ],
  ]
  var currentHeader = sheet.getRange('A1:Y1').getValues()

  if (currentHeader[0][0] == '') {
    sheet.getRange('A1:Y1').setValues(header)
  }

  var row = []
  row.push(header)

  var temp = []

  var yday = Utilities.formatDate(
    new Date(new Date() - 1000 * 60 * 60 * 24),
    'gmt',
    'dd/MM/yyyy'
  )
  temp.push(yday)

  for (item in detailedReport) {
    temp.push(detailedReport[item])
  }

  temp = temp.concat(totalReport)

  sheet.appendRow(temp)

  createGraphs()
  formatRange()
  transpose()
}
