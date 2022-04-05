import flask
import pathlib
import json
from SongxRevxHalf import songxrevxhalf
from IncomexRevxHalf import incomexrevxhalf
from SourcexRevxHalf import sourcexrevxhalf
from SongxIncomexRevxHalf import songxincomexrevxhalf
from ThirdPartyxSongxRevxHalf import thirdpartyxsongxrevxhalf
from ThirdPartyxIncomexRevxHalf import thirdpartyxincomexrevxhalf
from ThirdPartyxSourcexRevxHalf import thirdpartyxsourcexrevxhalf
from ArtistxRevxHalf import artistxrevxhalf
from Summary_Outputs import summary
from DCF import dcf
from CombinedOutputs import combined_outputs
from Songvest import songvest, data_periods
from Songvest2 import songvest2
from StandardOutputsx3rdParty import findthirdparties, thirdpartystandard
from SongvestOutputsx3rdParty import thirdpartysongvest
from ThirdPartyxArtistxRevxHalf import thirdpartyxartistxrevxhalf
from SongxTerritoryxRevxHalf import songxterritoryxrevxhalf
from databook import databook
from TerritoryxRevxHalf import territoryxrevxhalf
from CountryxRevxHalf import countryxrevxhalf
from RegionalBlockxRevxHalf import regionalblockxrevxhalf
from FullTerritory import fullterritoryxrevxhalf
from IncomeSource import incomesourcexrevxhalf
from SyncIncomeSource import syncincomesourcexrevxhalf
from SongID import songIDxrevxhalf
from SongDateIDISWCISRC import songdateIDISWCISRCxrevxhalf
from SongDateISWC import songdateISWCxrevxhalf
from SongDateISWCISRC import songdateISWCISRCxrevxhalf
from SongDateArtistIDISWCISRC import songdateartistIDISWCISRCxrevxhalf
from IncomeIIxRevxHalf import incomeIIxrevxhalf
from FullIncome import fullincomexrevxhalf
from SongxNetGross import songxnetgross
from SongIDIncomeSongShare import songIDincomesongshare
from SongIDIncomePayeePercentReceived import songIDincomepayeepercent
from SongIDISWCISRCIncomeSongShare import songIDISWCISRCincomesongshare
from SongIDISWCIncomeSongShare import songIDISWCincomesongshare
from SongIDISWCISRCIncomePayeePercentReceived import songIDISWCISRCincomepayeepercent
from SongIDISWCIncomePayeePercentReceived import songIDISWCincomepayeepercent
from SummaryNetGross import summarynetgross
from SongPlaysRevenue import songxplaysxrevenue
from IncomeTypePlaysRevenue import incomexplaysxrevenue
from SourcePlaysRevenue import sourcexplaysxrevenue
from PlaysxRevxHalf import playsxrevxhalf
from SongxIncomexPlaysxRevenue import songxincomexplaysxrevenue
from SongxSourcexPlaysxRevenue import songxsourcexplaysxrevenue
from SourceChainTotal import sourcechaintotal
from SourceChainxRevxHalf import sourcechainxrevxhalf
from SourceChainLines import sourcechainlines
from SourceChainxLinesxHalf import sourcechainxlinesxhalf
from IncomeCountryPayeePercent import incomecountrypayeepercent
from IncomeRegionalBlockPayeePercent import incomeregionalblockpayeepercent
from IncomeFullTerritoryPayeePercent import incomefullterritorypayeepercent
from IncomeCountrySongShare import incomecountrysongshare
from IncomeRegionalBlockSongShare import incomeregionalblocksongshare
from IncomeFullTerritorySongShare import incomefullterritorysongshare
from IncomeCountryPayeePercentRev import incomecountrypayeepercentrev
from IncomeRegionalBlockPayeePercentRev import incomeregionalblockpayeepercentrev
from IncomeCountrySongShareRev import incomecountrysongsharerev
from IncomeRegionalBlockSongShareRev import incomeregionalblocksongsharerev
from SongISWCISRC import songISWCISRC
from SongISWC import songISWC
from ThirdPartySongComposer import thirdpartysongcomposer
from USAStatusxRevxHalf import USAxrevxhalf
from SongxUSAStatusxRevxHalf import songxUSAxrevxhalf
from SongxPROBonusCredits import songxPRObonuscreditsxhalf
from SongxPROBonusDollars import songxPRObonusdollarsxhalf
from SongSpotifyID import songspotifyID
from SongAppleID import songappleID
from SongGeniusID import songgeniusID
from NullEventDateLines import nulleventdatelines
from NullEventDateRevenue import nulleventdaterev
from TempoRegionxRevxHalf import temporegionxrevxhalf
from TempoCountryxRevxHalf import tempocountryxrevxhalf
from ConfigurationxRevxHalf import configurationxrevxhalf
from IncomexPROBonusCredits import incomexPRObonuscreditsxhalf
from IncomexPROBonusDollars import incomexPRObonusdollarsxhalf
from SourcexPROBonusCredits import sourcexPRObonuscreditsxhalf
from SourcexPROBonusDollars import sourcexPRObonusdollarsxhalf
from PROBonusCreditsVDollars import creditsVdollars
from SongDatexRevxHalf import songdatexrevxhalf
from HipSongs import hipsongs
from HipIncome import hipincome
from HipSources import hipsources
from HipSourceChains import hipsourcechains
from HipTerritories import hipterritories
from HipNetGross import hipnetgross
from HipRates import hiprates
from HipRatesTerritories import hipratesterritories
from HipSongs2 import hipsongs2
from SyncxSourcexRevxHalf import syncxsourcexrevxhalf
from ReleaseYearxRevxHalf import releaseyearxrevxhalf
from SongxContractIDxRevxHalf import songxcontractidxrevxhalf
from HipSyncDetail import hipsyncdetail
from flask import request, jsonify
from flask_cors import cross_origin


local = False

app = flask.Flask(__name__)
def full_path(filename):
  if local:
    return f".\\output_files\\{filename}"
  else:
    return f"./output_files/{filename}"


@app.route('/songxrevxhalf', methods=['POST'])
@cross_origin()
def home():
  database = request.form.get('database')
  filename = 'SongxRevxHalf - {}.xlsx'.format(database[:-25])
  songxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomexrevxhalf', methods=['POST'])
@cross_origin()
def home1():
  database = request.form.get('database')
  filename = 'IncomexRevxHalf - {}.xlsx'.format(database[:-25])
  incomexrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/sourcexrevxhalf', methods=['POST'])
@cross_origin()
def home2():
  database = request.form.get('database')
  filename = 'SourcexRevxHalf - {}.xlsx'.format(database[:-25])
  sourcexrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songxincomexrevxhalf', methods=['POST'])
@cross_origin()
def home3():
  database = request.form.get('database')
  filename = 'SongxIncomexRevxHalf - {}.xlsx'.format(database[:-25])
  songxincomexrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/thirdpartyxsongxrevxhalf', methods=['POST'])
@cross_origin()
def home4():
  database = request.form.get('database')
  filename = 'ThirdPartyxSongxRevxHalf - {}.xlsx'.format(database[:-25])
  thirdpartyxsongxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/thirdpartyxincomexrevxhalf', methods=['POST'])
@cross_origin()
def home5():
  database = request.form.get('database')
  filename = 'ThirdPartyxIncomexRevxHalf - {}.xlsx'.format(database[:-25])
  thirdpartyxincomexrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/thirdpartyxsourcexrevxhalf', methods=['POST'])
@cross_origin()
def home6():
  database = request.form.get('database')
  filename = 'ThirdPartyxSourcexRevxHalf - {}.xlsx'.format(database[:-25])
  thirdpartyxsourcexrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/artistxrevxhalf', methods=['POST'])
@cross_origin()
def home7():
  database = request.form.get('database')
  filename = 'ArtistxRevxHalf - {}.xlsx'.format(database[:-25])
  artistxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/summaryoutputs', methods=['POST'])
@cross_origin()
def home8():
  database = request.form.get('database')
  filename = 'Summary - {}.xlsx'.format(database[:-25])
  summary(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/dcf', methods=['POST'])
@cross_origin()
def home9():
  database = request.form.get('database')
  growth_rates_string = request.form.get('growth_rates_string')
  growth_rates = json.loads(growth_rates_string)
  incremental_sync_string = request.form.get('incremental_sync_string')
  incremental_sync = json.loads(incremental_sync_string)
  discount_rate = request.form.get('discount_rate')
  tv_multiple = request.form.get('tv_multiple')
  tax_rate = request.form.get('tax_rate')
  initial_cost = request.form.get('initial_cost')
  filename = 'DCF - {}.xlsx'.format(database[:-25])
  dcf(database, full_path(filename), growth_rates, incremental_sync, discount_rate, tv_multiple, tax_rate, initial_cost)
  return jsonify({"file": filename})

@app.route('/masteroutput', methods=['POST'])
@cross_origin()
def home10():
  database = request.form.get('database')
  filename = 'Master Output - {}.xlsx'.format(database[:-25])
  combined_outputs(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/findperiod', methods=['POST'])
@cross_origin()
def home11():
  database = request.form.get('database')
  list = data_periods(database)
  return jsonify({"Data Periods": list})

@app.route('/songvest', methods=['POST'])
@cross_origin()
def home12():
  database = request.form.get('database')
  period = request.form.get('period')
  filename = 'Songvest - {}.xlsx'.format(database[:-25])
  songvest(database, full_path(filename), period)
  return jsonify({"file": filename})

@app.route('/songvest2', methods=['POST'])
@cross_origin()
def home13():
  database = request.form.get('database')
  filename = 'Songvest - {}.xlsx'.format(database[:-25])
  songvest2(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/findthirdparties', methods=['POST'])
@cross_origin()
def home14():
  database = request.form.get('database')
  list = findthirdparties(database)
  return jsonify({"Third Parties": list})

@app.route('/standardby3rdparty', methods=['POST'])
@cross_origin()
def home15():
  database = request.form.get('database')
  thirdparty = request.form.get('thirdparty')
  filename = '{} - {}.xlsx'.format(thirdparty,database[:-25])
  thirdpartystandard(database, full_path(filename), thirdparty)
  return jsonify({"file": filename})

@app.route('/songvestby3rdparty', methods=['POST'])
@cross_origin()
def home16():
  database = request.form.get('database')
  thirdparty = request.form.get('thirdparty')
  filename = '(Songvest) {} - {}.xlsx'.format(thirdparty,database[:-25])
  thirdpartysongvest(database, full_path(filename), thirdparty)
  return jsonify({"file": filename})

@app.route('/thirdpartyxartistxrevxhalf', methods=['POST'])
@cross_origin()
def home17():
  database = request.form.get('database')
  filename = 'ThirdPartyxArtistxRevxHalf - {}.xlsx'.format(database[:-25])
  thirdpartyxartistxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})


@app.route('/songxterritoryxrevxhalf', methods=['POST'])
@cross_origin()
def home19():
  database = request.form.get('database')
  filename = 'SongxTerritoryxRevxHalf - {}.xlsx'.format(database[:-25])
  songxterritoryxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/databook', methods=['POST'])
@cross_origin()
def home20():
  database = request.form.get('database')
  filename = 'Databook - {}.xlsx'.format(database[:-25])
  databook(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/countryxrevxhalf', methods=['POST'])
@cross_origin()
def home21():
  database = request.form.get('database')
  filename = 'Country x Rev x Half - {}.xlsx'.format(database[:-25])
  countryxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/regionalblockxrevxhalf', methods=['POST'])
@cross_origin()
def home22():
  database = request.form.get('database')
  filename = 'Regional Block x Rev x Half - {}.xlsx'.format(database[:-25])
  regionalblockxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/fullterritoryxrevxhalf', methods=['POST'])
@cross_origin()
def home23():
  database = request.form.get('database')
  filename = 'Full Territory x Rev x Half - {}.xlsx'.format(database[:-25])
  fullterritoryxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomesourcexrevxhalf', methods=['POST'])
@cross_origin()
def home24():
  database = request.form.get('database')
  filename = 'Income Type, Source x Rev x Half - {}.xlsx'.format(database[:-25])
  incomesourcexrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/syncincomesourcexrevxhalf', methods=['POST'])
@cross_origin()
def home25():
  database = request.form.get('database')
  filename = 'Sync Income Type, Source x Rev x Half - {}.xlsx'.format(database[:-25])
  syncincomesourcexrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songIDxrevxhalf', methods=['POST'])
@cross_origin()
def home26():
  database = request.form.get('database')
  filename = 'Song, ID x Rev x Half - {}.xlsx'.format(database[:-25])
  songIDxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songdateIDISWCISRCxrevxhalf', methods=['POST'])
@cross_origin()
def home27():
  database = request.form.get('database')
  filename = 'Song, Date, ID, ISWC, ISRC x Rev x Half - {}.xlsx'.format(database[:-25])
  songdateIDISWCISRCxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songdateISWCxrevxhalf', methods=['POST'])
@cross_origin()
def home28():
  database = request.form.get('database')
  filename = 'Song, Date, ISWC x Rev x Half - {}.xlsx'.format(database[:-25])
  songdateISWCxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songdateISWCISRCxrevxhalf', methods=['POST'])
@cross_origin()
def home29():
  database = request.form.get('database')
  filename = 'Song, Date, ISWC, ISRC x Rev x Half - {}.xlsx'.format(database[:-25])
  songdateISWCISRCxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songdateartistIDISWCISRCxrevxhalf', methods=['POST'])
@cross_origin()
def home30():
  database = request.form.get('database')
  filename = 'Song, Date, Artist, ID, ISWC, ISRC x Rev x Half - {}.xlsx'.format(database[:-25])
  songdateartistIDISWCISRCxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomeIIxrevxhalf', methods=['POST'])
@cross_origin()
def home31():
  database = request.form.get('database')
  filename = 'Income II x Rev x Half - {}.xlsx'.format(database[:-25])
  incomeIIxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/fullincomexrevxhalf', methods=['POST'])
@cross_origin()
def home32():
  database = request.form.get('database')
  filename = 'Income, Income II x Rev x Half - {}.xlsx'.format(database[:-25])
  fullincomexrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songxnetgross', methods=['POST'])
@cross_origin()
def home33():
  database = request.form.get('database')
  filename = 'Song x Net v Gross - {}.xlsx'.format(database[:-25])
  songxnetgross(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songIDincomesongshare', methods=['POST'])
@cross_origin()
def home34():
  database = request.form.get('database')
  filename = 'Song, ID, Income Type, Song Share - {}.xlsx'.format(database[:-25])
  songIDincomesongshare(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songIDincomepayeepercent', methods=['POST'])
@cross_origin()
def home35():
  database = request.form.get('database')
  filename = 'Song, ID, Income Type, Payee Percent - {}.xlsx'.format(database[:-25])
  songIDincomepayeepercent(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songIDISWCISRCincomesongshare', methods=['POST'])
@cross_origin()
def home36():
  database = request.form.get('database')
  filename = 'Song, ID, ISWC, ISRC, Income Type, Song Share - {}.xlsx'.format(database[:-25])
  songIDISWCISRCincomesongshare(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songIDISWCincomesongshare', methods=['POST'])
@cross_origin()
def home37():
  database = request.form.get('database')
  filename = 'Song, ID, ISWC, Income Type, Song Share - {}.xlsx'.format(database[:-25])
  songIDISWCincomesongshare(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songIDISWCISRCincomepayeepercentrecieved', methods=['POST'])
@cross_origin()
def home38():
  database = request.form.get('database')
  filename = 'Song, ID, ISWC, ISRC, Income Type, Payee Percent - {}.xlsx'.format(database[:-25])
  songIDISWCISRCincomepayeepercent(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songIDISWCincomepayeepercentrecieved', methods=['POST'])
@cross_origin()
def home39():
  database = request.form.get('database')
  filename = 'Song, ID, ISWC, Income Type, Payee Percent - {}.xlsx'.format(database[:-25])
  songIDISWCincomepayeepercent(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/summarynetgross', methods=['POST'])
@cross_origin()
def home40():
  database = request.form.get('database')
  filename = 'Net v Gross Summary - {}.xlsx'.format(database[:-25])
  summarynetgross(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songxplaysxrevenue', methods=['POST'])
@cross_origin()
def home41():
  database = request.form.get('database')
  filename = 'Song x Plays x Revenue - {}.xlsx'.format(database[:-25])
  songxplaysxrevenue(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/playsxrevxhalf', methods=['POST'])
@cross_origin()
def home42():
  database = request.form.get('database')
  filename = 'Plays x Rev x Half - {}.xlsx'.format(database[:-25])
  playsxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomexplaysxrevenue', methods=['POST'])
@cross_origin()
def home43():
  database = request.form.get('database')
  filename = 'Income Type x Plays x Revenue - {}.xlsx'.format(database[:-25])
  incomexplaysxrevenue(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/sourcexplaysxrevenue', methods=['POST'])
@cross_origin()
def home44():
  database = request.form.get('database')
  filename = 'Source x Plays x Revenue - {}.xlsx'.format(database[:-25])
  sourcexplaysxrevenue(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songxincomexplaysxrevenue', methods=['POST'])
@cross_origin()
def home45():
  database = request.form.get('database')
  filename = 'Song x Income x Plays x Revenue - {}.xlsx'.format(database[:-25])
  songxincomexplaysxrevenue(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songxsourcexplaysxrevenue', methods=['POST'])
@cross_origin()
def home46():
  database = request.form.get('database')
  filename = 'Song x Source x Plays x Revenue - {}.xlsx'.format(database[:-25])
  songxsourcexplaysxrevenue(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/sourcechaintotal', methods=['POST'])
@cross_origin()
def home48():
  database = request.form.get('database')
  filename = 'Source Chain x Total - {}.xlsx'.format(database[:-25])
  sourcechaintotal(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/sourcechainxrevxhalf', methods=['POST'])
@cross_origin()
def home49():
  database = request.form.get('database')
  filename = 'Source Chain x Rev x Half - {}.xlsx'.format(database[:-25])
  sourcechainxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/sourcechainlines', methods=['POST'])
@cross_origin()
def home50():
  database = request.form.get('database')
  filename = 'Source Chain x Lines - {}.xlsx'.format(database[:-25])
  sourcechainlines(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/sourcechainxlinesxhalf', methods=['POST'])
@cross_origin()
def home51():
  database = request.form.get('database')
  filename = 'Source Chain x Lines x Half - {}.xlsx'.format(database[:-25])
  sourcechainxlinesxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/territoryxrevxhalf', methods=['POST'])
@cross_origin()
def home52():
  database = request.form.get('database')
  filename = 'Territory x Rev x Half - {}.xlsx'.format(database[:-25])
  territoryxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomecountrypayeepercent', methods=['POST'])
@cross_origin()
def home53():
  database = request.form.get('database')
  filename = 'Income, Country, Payee Percent - {}.xlsx'.format(database[:-25])
  incomecountrypayeepercent(database, full_path(filename))
  return jsonify({"file": filename})


@app.route('/incomeregionalblockpayeepercent', methods=['POST'])
@cross_origin()
def home54():
  database = request.form.get('database')
  filename = 'Income, Regional Block, Payee Percent - {}.xlsx'.format(database[:-25])
  incomeregionalblockpayeepercent(database, full_path(filename))
  return jsonify({"file": filename})


@app.route('/incomefullterritorypayeepercent', methods=['POST'])
@cross_origin()
def home55():
  database = request.form.get('database')
  filename = 'Income, Full Territory, Payee Percent - {}.xlsx'.format(database[:-25])
  incomefullterritorypayeepercent(database, full_path(filename))
  return jsonify({"file": filename})


@app.route('/incomecountrysongshare', methods=['POST'])
@cross_origin()
def home56():
  database = request.form.get('database')
  filename = 'Income, Country, Song Share - {}.xlsx'.format(database[:-25])
  incomecountrysongshare(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomeregionalblocksongshare', methods=['POST'])
@cross_origin()
def home57():
  database = request.form.get('database')
  filename = 'Income, Regional Block, Song Share - {}.xlsx'.format(database[:-25])
  incomeregionalblocksongshare(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomefullterritorysongshare', methods=['POST'])
@cross_origin()
def home58():
  database = request.form.get('database')
  filename = 'Income, Regional Block, Region, Country, Song Share - {}.xlsx'.format(database[:-25])
  incomefullterritorysongshare(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomecountrypayeepercentrev', methods=['POST'])
@cross_origin()
def home59():
  database = request.form.get('database')
  filename = 'Income, Country, Payee Percent, Rev - {}.xlsx'.format(database[:-25])
  incomecountrypayeepercentrev(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomeregionalblockpayeepercentrev', methods=['POST'])
@cross_origin()
def home60():
  database = request.form.get('database')
  filename = 'Income, Regional Block, Payee Percent, Rev - {}.xlsx'.format(database[:-25])
  incomeregionalblockpayeepercentrev(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomecountrysongsharerev', methods=['POST'])
@cross_origin()
def home61():
  database = request.form.get('database')
  filename = 'Income, Country, Song Share, Rev - {}.xlsx'.format(database[:-25])
  incomecountrysongsharerev(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomeregionalblocksongsharerev', methods=['POST'])
@cross_origin()
def home62():
  database = request.form.get('database')
  filename = 'Income, Regional Block, Song Share, Rev - {}.xlsx'.format(database[:-25])
  incomeregionalblocksongsharerev(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songISWCISRC', methods=['POST'])
@cross_origin()
def home63():
  database = request.form.get('database')
  filename = 'Song, ISWC, ISRC - {}.xlsx'.format(database[:-25])
  songISWCISRC(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songISWC', methods=['POST'])
@cross_origin()
def home64():
  database = request.form.get('database')
  filename = 'Song, ISWC - {}.xlsx'.format(database[:-25])
  songISWC(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/thirdpartysongcomposer', methods=['POST'])
@cross_origin()
def home65():
  database = request.form.get('database')
  filename = 'Third Party,Song,Composer - {}.xlsx'.format(database[:-25])
  thirdpartysongcomposer(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/USAxrevxhalf', methods=['POST'])
@cross_origin()
def home66():
  database = request.form.get('database')
  filename = 'USA x Rev x Half - {}.xlsx'.format(database[:-25])
  USAxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songxUSAxrevxhalf', methods=['POST'])
@cross_origin()
def home67():
  database = request.form.get('database')
  filename = 'Song x USA x Rev x Half - {}.xlsx'.format(database[:-25])
  songxUSAxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songxPRObonuscredits', methods=['POST'])
@cross_origin()
def home68():
  database = request.form.get('database')
  filename = 'Song x PRO Bonus Credits x Half - {}.xlsx'.format(database[:-25])
  songxPRObonuscreditsxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songxPRObonusdollars', methods=['POST'])
@cross_origin()
def home69():
  database = request.form.get('database')
  filename = 'Song x PRO Bonus Dollars x Half - {}.xlsx'.format(database[:-25])
  songxPRObonusdollarsxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songspotifyID', methods=['POST'])
@cross_origin()
def home70():
  database = request.form.get('database')
  filename = 'Song, Spotify ID x Rev x Half - {}.xlsx'.format(database[:-25])
  songspotifyID(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songappleID', methods=['POST'])
@cross_origin()
def home71():
  database = request.form.get('database')
  filename = 'Song, Apple ID x Rev x Half - {}.xlsx'.format(database[:-25])
  songappleID(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songgeniusID', methods=['POST'])
@cross_origin()
def home72():
  database = request.form.get('database')
  filename = 'Song, Genius ID x Rev x Half - {}.xlsx'.format(database[:-25])
  songgeniusID(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/nulleventdatelines', methods=['POST'])
@cross_origin()
def home73():
  database = request.form.get('database')
  filename = 'Null Event Date Lines - {}.xlsx'.format(database[:-25])
  nulleventdatelines(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/nulleventdaterev', methods=['POST'])
@cross_origin()
def home74():
  database = request.form.get('database')
  filename = 'Null Event Date Rev - {}.xlsx'.format(database[:-25])
  nulleventdaterev(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/temporegionxrevxhalf', methods=['POST'])
@cross_origin()
def home75():
  database = request.form.get('database')
  filename = 'Region x Rev x Half - {}.xlsx'.format(database[:-25])
  temporegionxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/tempocountryxrevxhalf', methods=['POST'])
@cross_origin()
def home76():
  database = request.form.get('database')
  filename = 'Country x Rev x Half - {}.xlsx'.format(database[:-25])
  tempocountryxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/configurationxrevxhalf', methods=['POST'])
@cross_origin()
def home77():
  database = request.form.get('database')
  filename = 'Configuration x Rev x Half - {}.xlsx'.format(database[:-25])
  configurationxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomexPRObonuscredits', methods=['POST'])
@cross_origin()
def home78():
  database = request.form.get('database')
  filename = 'Income x PRO Bonus Credits x Half - {}.xlsx'.format(database[:-25])
  incomexPRObonuscreditsxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/incomexPRObonusdollars', methods=['POST'])
@cross_origin()
def home79():
  database = request.form.get('database')
  filename = 'Income x PRO Bonus Dollars x Half - {}.xlsx'.format(database[:-25])
  incomexPRObonusdollarsxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/sourcexPRObonuscredits', methods=['POST'])
@cross_origin()
def home80():
  database = request.form.get('database')
  filename = 'Source x PRO Bonus Credits x Half - {}.xlsx'.format(database[:-25])
  sourcexPRObonuscreditsxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/sourcexPRObonusdollars', methods=['POST'])
@cross_origin()
def home81():
  database = request.form.get('database')
  filename = 'Source x PRO Bonus Dollars x Half - {}.xlsx'.format(database[:-25])
  sourcexPRObonusdollarsxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/creditsVdollars', methods=['POST'])
@cross_origin()
def home82():
  database = request.form.get('database')
  filename = 'Credits V Dollars - {}.xlsx'.format(database[:-25])
  creditsVdollars(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songdatexrevxhalf', methods=['POST'])
@cross_origin()
def home83():
  database = request.form.get('database')
  filename = 'Song,Date x Rev x Half - {}.xlsx'.format(database[:-25])
  songdatexrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/hipsongs', methods=['POST'])
@cross_origin()
def home84():
  database = request.form.get('database')
  filename = 'Hipgnosis Songs - {}.xlsx'.format(database[:-25])
  hipsongs(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/hipincome', methods=['POST'])
@cross_origin()
def home85():
  database = request.form.get('database')
  filename = 'Hipgnosis Income - {}.xlsx'.format(database[:-25])
  hipincome(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/hipsources', methods=['POST'])
@cross_origin()
def home86():
  database = request.form.get('database')
  filename = 'Hipgnosis Sources - {}.xlsx'.format(database[:-25])
  hipsources(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/hipsourcechains', methods=['POST'])
@cross_origin()
def home87():
  database = request.form.get('database')
  filename = 'Hipgnosis Source Chains - {}.xlsx'.format(database[:-25])
  hipsourcechains(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/hipterritories', methods=['POST'])
@cross_origin()
def home88():
  database = request.form.get('database')
  filename = 'Hipgnosis Territories - {}.xlsx'.format(database[:-25])
  hipterritories(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/hipnetgross', methods=['POST'])
@cross_origin()
def home89():
  database = request.form.get('database')
  filename = 'Hipgnosis Net Gross - {}.xlsx'.format(database[:-25])
  hipnetgross(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/hiprates', methods=['POST'])
@cross_origin()
def home90():
  database = request.form.get('database')
  filename = 'Hipgnosis Rates - {}.xlsx'.format(database[:-25])
  hiprates(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/hipratesterritories', methods=['POST'])
@cross_origin()
def home91():
  database = request.form.get('database')
  filename = 'Hipgnosis Rates Territories - {}.xlsx'.format(database[:-25])
  hipratesterritories(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/hipsongs2', methods=['POST'])
@cross_origin()
def home92():
  database = request.form.get('database')
  filename = 'Hipgnosis Songs - {}.xlsx'.format(database[:-25])
  hipsongs2(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/syncxsourcexrevxhalf', methods=['POST'])
@cross_origin()
def home93():
  database = request.form.get('database')
  filename = 'Sync x Source x Rev x Half - {}.xlsx'.format(database[:-25])
  syncxsourcexrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/releaseyearxrevxhalf', methods=['POST'])
@cross_origin()
def home94():
  database = request.form.get('database')
  filename = 'Release Year x Rev x Half - {}.xlsx'.format(database[:-25])
  releaseyearxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/songxcontractidxrevxhalf', methods=['POST'])
@cross_origin()
def home95():
  database = request.form.get('database')
  filename = 'Song x Contract ID x Rev x Half - {}.xlsx'.format(database[:-25])
  songxcontractidxrevxhalf(database, full_path(filename))
  return jsonify({"file": filename})

@app.route('/hipsyncdetail', methods=['POST'])
@cross_origin()
def home96():
  database = request.form.get('database')
  filename = 'Hipgnosis Sync Detail - {}.xlsx'.format(database[:-25])
  hipsyncdetail(database, full_path(filename))
  return jsonify({"file": filename})


app.run(port=5050)






