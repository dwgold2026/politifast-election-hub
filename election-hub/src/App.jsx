import { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const S = [
{s:"AL",n:"Alabama",filing:"Jan 23, 2026",primary:"May 19, 2026",runoff:"Jun 16, 2026",general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:true,govNote:"Tuberville vacating Senate to run",senate:true,senOpen:true,senNote:"Open — Tuberville running for Gov",house:true,
otherOffices:["Attorney General","Secretary of State","Full Legislature"],
centralDb:{name:"AL Secretary of State",url:"https://sos.alabama.gov/alabama-votes",covers:"Federal, statewide, legislature",format:"Online search",bulk:false},
localType:"County Probate Judge",localCount:67,localDir:"https://www.sos.alabama.gov/city-county-lookup/probate-judges",
counties:[
["Autauga","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Baldwin","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Barbour","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Bibb","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Blount","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Bullock","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Butler","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Calhoun","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Chambers","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Cherokee","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Chilton","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Choctaw","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Clarke","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Clay","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Cleburne","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Coffee","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Colbert","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Conecuh","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Coosa","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Covington","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Crenshaw","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Cullman","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Dale","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Dallas","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["DeKalb","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Elmore","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Escambia","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Etowah","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Fayette","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Franklin","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Geneva","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Greene","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Hale","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Henry","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Houston","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Jackson","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Jefferson","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Lamar","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Lauderdale","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Lawrence","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Lee","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Limestone","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Lowndes","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Macon","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Madison","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Marengo","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Marion","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Marshall","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Mobile","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Monroe","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Montgomery","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Morgan","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Perry","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Pickens","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Pike","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Randolph","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Russell","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Shelby","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["St. Clair","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Sumter","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Talladega","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Tallapoosa","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Tuscaloosa","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Walker","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Washington","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Wilcox","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],["Winston","https://www.sos.alabama.gov/city-county-lookup/probate-judges"],
],
gap:"Municipal, school board, special district candidates NOT in state database.",
schoolNote:"Data held by County Superintendent of Education. Partisan school board races (unusual).",
notes:"Tommy Tuberville vacating Senate to run for Governor creates rare dual open-seat cycle."},
{s:"AK",n:"Alaska",filing:"Jun 1, 2026",primary:"Aug 18, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:true,govNote:"Term-limited",senate:true,senOpen:false,senNote:"Sullivan (R)",house:true,
otherOffices:["Top-four primary + RCV general","Legislature"],
centralDb:{name:"AK Division of Elections",url:"https://www.elections.alaska.gov/candidates/",covers:"Federal, statewide, legislature",format:"Online search",bulk:false},
localType:"Borough Clerk",localCount:19,localDir:"https://www.elections.alaska.gov/",
counties:[
["Aleutians East","https://www.aleutianseast.org/"],["Anchorage","https://www.muni.org/Departments/Assembly/Clerk/Elections/Pages/default.aspx"],["Bristol Bay","https://www.bristolbayboroughak.us/"],["Denali","https://www.denaliborough.org/election"],["Fairbanks North Star","https://www.fnsb.gov/269/Election-Information"],["Haines","https://www.hainesalaska.gov/administration/page/borough-elections"],["Juneau","https://juneau.org/clerk/elections"],["Kenai Peninsula","https://www.kpb.us/local-governance-and-permitting/voting-elections/voting-elections-overview"],["Ketchikan Gateway","https://www.kgbak.us/172/Elections"],["Kodiak Island","https://www.kodiakak.us/234/Elections"],["Lake and Peninsula","https://lakeandpen.com/"],["Matanuska-Susitna","https://matsugov.us/elections"],["North Slope","https://www.north-slope.org/departments/assembly-clerk/clerks-office/"],["Northwest Arctic","https://www.nwabor.org/"],["Petersburg","https://www.petersburgak.gov/"],["Sitka","https://www.cityofsitka.com/RegularMunicipalElection"],["Skagway","https://www.skagway.org/clerksoffice/page/election-information"],["Wrangell","https://www.wrangell.com/administration/borough-local-election-information"],["Yakutat","https://www.yakutatak.us/borough-assembly/page/voting-elections"],
],
gap:"Local/school board not centralized, but only 19 boroughs — very manageable.",
schoolNote:"Data held by borough or city clerk.",
notes:"Top-four primary with ranked-choice general. Only 19 organized boroughs."},
{s:"AZ",n:"Arizona",filing:"Mar 23, 2026",primary:"Jul 21, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:false,govNote:"",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Full Legislature"],
centralDb:{name:"AZ Secretary of State",url:"https://azsos.gov/elections/candidate-information",covers:"Federal, statewide, legislature",format:"Online search",bulk:false},
localType:"County Recorder",localCount:15,localDir:"https://azsos.gov/elections/voting-election/contact-information-county-election-officials",
counties:[
["Apache","https://www.co.apache.az.us/elections/"],["Cochise","https://www.cochise.az.gov/509/Voter-Information"],["Coconino","https://www.coconino.az.gov/195/Elections"],["Gila","https://www.gilacountyaz.gov/government/elections/"],["Graham","https://www.graham.az.gov/476/Elections"],["Greenlee","https://greenlee.az.gov/departments/elections"],["La Paz","https://www.co.la-paz.az.us/162/Elections"],["Maricopa","https://elections.maricopa.gov/"],["Mohave","https://www.mohave.gov/departments/elections/"],["Navajo","https://www.navajocountyaz.gov/223/Elections"],["Pima","https://www.pima.gov/394/Elections"],["Pinal","https://www.pinal.gov/258/Elections"],["Santa Cruz","https://www.santacruzcountyaz.gov/173/Elections"],["Yavapai","https://www.yavapaivotes.gov/"],["Yuma","https://www.yumacountyaz.gov/government/election-services"],
],
gap:"County, municipal, school board require county-level lookups.",
schoolNote:"Data held by County School Superintendent. Maricopa: schoolsup.org.",
notes:"No Senate race. Maricopa County (60% of pop) has good online tools."},
{s:"AR",n:"Arkansas",filing:"Nov 12, 2025",primary:"Mar 3, 2026",runoff:"Mar 31, 2026",general:"Nov 3, 2026",status:"voted",
gov:true,govOpen:false,govNote:"Sanders (R) unopposed",senate:true,senOpen:false,senNote:"Cotton (R) renominated",house:true,
otherOffices:["Lt. Governor","Attorney General","Secretary of State","Full Legislature"],
centralDb:{name:"AR SoS Candidate Search",url:"https://candidates.arkansas.gov/",covers:"Federal, statewide, legislature, some county",format:"Searchable database",bulk:false},
localType:"County Clerk",localCount:75,localDir:"https://www.sos.arkansas.gov/uploads/elections/ARCountyClerks.pdf",
counties:[
["Arkansas","https://www.sos.arkansas.gov/elections"],["Ashley","https://www.sos.arkansas.gov/elections"],["Baxter","https://www.sos.arkansas.gov/elections"],["Benton","https://www.sos.arkansas.gov/elections"],["Boone","https://www.sos.arkansas.gov/elections"],["Bradley","https://www.sos.arkansas.gov/elections"],["Calhoun","https://www.sos.arkansas.gov/elections"],["Carroll","https://www.sos.arkansas.gov/elections"],["Chicot","https://www.sos.arkansas.gov/elections"],["Clark","https://www.sos.arkansas.gov/elections"],["Clay","https://www.sos.arkansas.gov/elections"],["Cleburne","https://www.sos.arkansas.gov/elections"],["Cleveland","https://www.sos.arkansas.gov/elections"],["Columbia","https://www.sos.arkansas.gov/elections"],["Conway","https://www.sos.arkansas.gov/elections"],["Craighead","https://www.sos.arkansas.gov/elections"],["Crawford","https://www.sos.arkansas.gov/elections"],["Crittenden","https://www.sos.arkansas.gov/elections"],["Cross","https://www.sos.arkansas.gov/elections"],["Dallas","https://www.sos.arkansas.gov/elections"],["Desha","https://www.sos.arkansas.gov/elections"],["Drew","https://www.sos.arkansas.gov/elections"],["Faulkner","https://www.sos.arkansas.gov/elections"],["Franklin","https://www.sos.arkansas.gov/elections"],["Fulton","https://www.sos.arkansas.gov/elections"],["Garland","https://www.sos.arkansas.gov/elections"],["Grant","https://www.sos.arkansas.gov/elections"],["Greene","https://www.sos.arkansas.gov/elections"],["Hempstead","https://www.sos.arkansas.gov/elections"],["Hot Spring","https://www.sos.arkansas.gov/elections"],["Howard","https://www.sos.arkansas.gov/elections"],["Independence","https://www.sos.arkansas.gov/elections"],["Izard","https://www.sos.arkansas.gov/elections"],["Jackson","https://www.sos.arkansas.gov/elections"],["Jefferson","https://www.sos.arkansas.gov/elections"],["Johnson","https://www.sos.arkansas.gov/elections"],["Lafayette","https://www.sos.arkansas.gov/elections"],["Lawrence","https://www.sos.arkansas.gov/elections"],["Lee","https://www.sos.arkansas.gov/elections"],["Lincoln","https://www.sos.arkansas.gov/elections"],["Little River","https://www.sos.arkansas.gov/elections"],["Logan","https://www.sos.arkansas.gov/elections"],["Lonoke","https://www.sos.arkansas.gov/elections"],["Madison","https://www.sos.arkansas.gov/elections"],["Marion","https://www.sos.arkansas.gov/elections"],["Miller","https://www.sos.arkansas.gov/elections"],["Mississippi","https://www.sos.arkansas.gov/elections"],["Monroe","https://www.sos.arkansas.gov/elections"],["Montgomery","https://www.sos.arkansas.gov/elections"],["Nevada","https://www.sos.arkansas.gov/elections"],["Newton","https://www.sos.arkansas.gov/elections"],["Ouachita","https://www.sos.arkansas.gov/elections"],["Perry","https://www.sos.arkansas.gov/elections"],["Phillips","https://www.sos.arkansas.gov/elections"],["Pike","https://www.sos.arkansas.gov/elections"],["Poinsett","https://www.sos.arkansas.gov/elections"],["Polk","https://www.sos.arkansas.gov/elections"],["Pope","https://www.sos.arkansas.gov/elections"],["Prairie","https://www.sos.arkansas.gov/elections"],["Pulaski","https://www.sos.arkansas.gov/elections"],["Randolph","https://www.sos.arkansas.gov/elections"],["Saline","https://www.sos.arkansas.gov/elections"],["Scott","https://www.sos.arkansas.gov/elections"],["Searcy","https://www.sos.arkansas.gov/elections"],["Sebastian","https://www.sos.arkansas.gov/elections"],["Sevier","https://www.sos.arkansas.gov/elections"],["Sharp","https://www.sos.arkansas.gov/elections"],["St. Francis","https://www.sos.arkansas.gov/elections"],["Stone","https://www.sos.arkansas.gov/elections"],["Union","https://www.sos.arkansas.gov/elections"],["Van Buren","https://www.sos.arkansas.gov/elections"],["Washington","https://www.sos.arkansas.gov/elections"],["White","https://www.sos.arkansas.gov/elections"],["Woodruff","https://www.sos.arkansas.gov/elections"],["Yell","https://www.sos.arkansas.gov/elections"],
],
gap:"Smaller local and school board races may not appear in state database.",
schoolNote:"Data held by county clerks.",
notes:"Primary held March 3. Runoff March 31 if needed."},
{s:"CA",n:"California",filing:"Mar 6, 2026",primary:"Jun 2, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:true,govNote:"Newsom term-limited",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Controller","Insurance Commissioner","Legislature"],
centralDb:{name:"Cal-Access",url:"https://cal-access.sos.ca.gov/Campaign/Candidates/",covers:"Federal, statewide, legislature",format:"Searchable database + API",bulk:true},
localType:"County Registrar of Voters",localCount:58,localDir:"https://www.sos.ca.gov/elections/voting-resources/county-elections-offices",
counties:[["Alameda","https://www.acvote.org/"],["Alpine","https://www.alpinecountyca.gov/196/Elections"],["Amador","https://www.amadorgov.org/government/elections"],["Butte","https://www.buttecounty.net/elections"],["Calaveras","https://elections.calaverasgov.us/"],["Colusa","https://www.countyofcolusa.org/index.aspx?NID=112"],["Contra Costa","https://www.cocovote.us/"],["Del Norte","https://www.co.del-norte.ca.us/departments/elections"],["El Dorado","https://www.edcgov.us/Government/Elections"],["Fresno","https://www.co.fresno.ca.us/departments/county-clerk-registrar-of-voters"],["Glenn","https://www.countyofglenn.net/dept/elections/welcome"],["Humboldt","https://humboldtgov.org/890/Elections"],["Imperial","https://www.co.imperial.ca.us/registrarofvoters/"],["Inyo","https://www.inyocounty.us/services/elections"],["Kern","https://www.kernvote.com/"],["Kings","https://www.countyofkings.com/departments/administration/elections"],["Lake","https://www.lakecountyca.gov/Government/Directory/Registrar_of_Voters.htm"],["Lassen","https://www.lassencounty.org/dept/county-clerk-recorder/elections"],["Los Angeles","https://lavote.gov/"],["Madera","https://votemadera.com/"],["Marin","https://www.marinvotes.org/"],["Mariposa","https://www.mariposacounty.org/96/Elections"],["Mendocino","https://www.mendocinocounty.org/government/assessor-county-clerk-recorder-elections/elections"],["Merced","https://www.co.merced.ca.us/elections"],["Modoc","https://www.co.modoc.ca.us/departments/elections"],["Mono","https://monocounty.ca.gov/elections"],["Monterey","https://www.montereycountyelections.us/"],["Napa","https://www.countyofnapa.org/152/Elections"],["Nevada","https://www.mynevadacounty.com/2364/Elections-Voting"],["Orange","https://www.ocvote.gov/"],["Placer","https://www.placerelections.com/"],["Plumas","https://www.plumascounty.us/129/Elections"],["Riverside","https://www.voteinfo.net/"],["Sacramento","https://elections.saccounty.gov/"],["San Benito","https://www.sbcvote.us/"],["San Bernardino","https://www.sbcountyelections.com/"],["San Diego","https://www.sdvote.com/"],["San Francisco","https://sfelections.sfgov.org/"],["San Joaquin","https://www.sjcrov.org/"],["San Luis Obispo","https://www.slovote.com/"],["San Mateo","https://www.smcacre.org/elections"],["Santa Barbara","https://countyofsb.org/care/elections"],["Santa Clara","https://www.sccvote.org/"],["Santa Cruz","https://www.votescount.us/"],["Shasta","https://www.elections.co.shasta.ca.us/"],["Sierra","https://www.sierracounty.ca.gov/171/Elections"],["Siskiyou","https://www.co.siskiyou.ca.us/elections"],["Solano","https://www.solanocounty.com/depts/rov/default.asp"],["Sonoma","https://sonomacounty.ca.gov/administrative-support-and-fiscal-services/clerk-recorder-assessor-registrar-of-voters/registrar-of-voters"],["Stanislaus","https://www.stanvote.com/"],["Sutter","https://www.suttercounty.org/government/county-departments/elections"],["Tehama","https://www.co.tehama.ca.us/government/county-departments-offices/county-clerk"],["Trinity","https://www.trinitycounty.org/Elections"],["Tulare","https://tularecounty.ca.gov/registrarofvoters/"],["Tuolumne","https://www.tuolumnecounty.ca.gov/194/Elections"],["Ventura","https://recorder.countyofventura.org/elections/"],["Yolo","https://www.yoloelections.org/"],["Yuba","https://www.yuba.org/departments/elections/index.php"]],
gap:"ALL local data decentralized across 58 counties. No state aggregation for local races.",
schoolNote:"Data held by county registrar or county superintendent of schools.",
notes:"Top-two primary. Newsom term-limited. Largest number of local jurisdictions."},
{s:"CO",n:"Colorado",filing:"Mar 18, 2026",primary:"Jun 30, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:true,govNote:"Polis term-limited",senate:true,senOpen:false,senNote:"Hickenlooper (D)",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Legislature"],
centralDb:{name:"TRACER Campaign Finance",url:"https://tracer.sos.colorado.gov/PublicSite/SearchPages/CandidateSearch.aspx",covers:"Federal, statewide, legislature",format:"Searchable database",bulk:true},
localType:"County Clerk and Recorder",localCount:64,localDir:"https://www.sos.state.co.us/pubs/elections/Resources/CountyElectionOffices.html",
counties:[
["Adams","https://www.adamscountycoelections.gov/"],
["Alamosa","https://www.alamosacounty.org/166/Elections-Department"],
["Arapahoe","https://www.arapahoevotes.gov"],
["Archuleta","https://www.archuletacounty.gov/county-government/elected-offices/clerk-recorder/elections/"],
["Baca","https://www.bacacountyco.gov/government/clerk-and-recorder/elections/"],
["Bent","https://www.bentcounty.net/government/clerk.php"],
["Boulder","https://bouldercountyvotes.gov"],
["Broomfield","https://www.broomfield.org/index.aspx?nid=191"],
["Chaffee","https://chaffeeclerk.colorado.gov/elections"],
["Cheyenne","https://www.co.cheyenne.co.us/departments/clerk_recorder.html"],
["Clear Creek","https://www.clearcreekcounty.us/104/Clerk-Recorder"],
["Conejos","https://conejoscounty.colorado.gov/elected-officials/clerk-and-recorder"],
["Costilla","https://www.costillacounty.gov/clerk-recorder"],
["Crowley","https://crowleycounty.colorado.gov/government-offices/clerk-and-recorder-elections"],
["Custer","https://www.custercounty-co.gov/elections"],
["Delta","https://www.deltacounty.com/325/Elections"],
["Denver","https://www.denvergov.org/content/denvergov/en/denver-elections-divison.html"],
["Dolores","https://dolocnty.colorado.gov/clerk-and-recorder"],
["Douglas","https://www.douglasvotes.com/"],
["Eagle","https://www.eaglecounty.us/vote"],
["El Paso","https://clerkandrecorder.elpasoco.com"],
["Elbert","https://www.elbertcounty-co.gov/290/Elections"],
["Fremont","https://www.fremontcountyelectionsco.gov/"],
["Garfield","https://www.garfieldcountyco.gov/clerk-recorder/elections/"],
["Gilpin","https://gilpincounty.colorado.gov/elected-officials/clerk-recorder"],
["Grand","https://www.co.grand.co.us/147/Elections"],
["Gunnison","https://www.gunnisoncounty.org/143/Clerk-Recorders-Office"],
["Hinsdale","https://hinsdalecounty.colorado.gov/hinsdale-county-clerk"],
["Huerfano","https://huerfano.us/departments/clerk-recorder/"],
["Jackson","https://jacksoncounty.colorado.gov/clerk"],
["Jefferson","https://jeffco.us/elections/"],
["Kiowa","https://kiowacounty-colorado.com/kiowa_county_clerk__recorder.htm"],
["Kit Carson","https://kitcarsoncounty.colorado.gov/clerk-recorder/elections"],
["La Plata","https://www.co.laplata.co.us/services/elections/index.php"],
["Lake","https://www.lakecountyco.gov/354/Elections"],
["Larimer","https://www.larimer.org/clerk/"],
["Las Animas","https://lasanimascounty.colorado.gov/elected-officials/clerk-and-recorder"],
["Lincoln","https://lincolncounty.colorado.gov/elections-legislators"],
["Logan","https://www.logancountyco.gov/186/Elections-Voter-Registration-Information"],
["Mesa","https://www.mesacounty.us/departments-and-services/clerk-and-recorder/elections"],
["Mineral","https://www.mineralcountycolorado.com/clerk-recorder/page/elections"],
["Moffat","https://moffatcounty.colorado.gov/government/elected-officials/county-clerk-and-recorder/elections"],
["Montezuma","https://montezumacounty.org/elections-office/"],
["Montrose","https://www.montrosecounty.net/73/Elections"],
["Morgan","https://morgancounty.colorado.gov/elections-department"],
["Otero","https://oterocounty.colorado.gov/elected-officials/clerk-recorder/elections"],
["Ouray","https://ouraycountyco.gov/328/Election-Information"],
["Park","https://www.parkco.us/index.aspx?nid=72"],
["Phillips","https://phillipscounty.colorado.gov/elections"],
["Pitkin","https://www.pitkinvotes.com/"],
["Prowers","https://www.prowersco.gov/departments/elections"],
["Pueblo","https://pueblovotes.com/"],
["Rio Blanco","https://www.rbc.us/176/Clerk-Recorder"],
["Rio Grande","https://riograndecounty.colorado.gov/how-do-i/elections/voter-information"],
["Routt","https://www.co.routt.co.us/221/Elections"],
["Saguache","https://www.colorado.gov/saguachecoclerk"],
["San Juan","https://sanjuancounty.colorado.gov/elections"],
["San Miguel","https://www.sanmiguelcountyco.gov/163/Clerk-Recorder"],
["Sedgwick","https://sedgwickcounty.colorado.gov/elected-officials/clerk-recorder"],
["Summit","https://www.summitcountyco.gov/services/clerk_recorder/elections/index.php"],
["Teller","https://www.tellercounty.gov/Elections"],
["Washington","https://washingtoncounty.colorado.gov/departments/elections"],
["Weld","https://www.weldgov.com/Government/Departments/Clerk-and-Recorder/Elections-Department"],
["Yuma","https://yumacounty.net/clerk-and-recorder/"],
],
gap:"County, municipal, school board, special district not in TRACER.",
schoolNote:"Data held by county clerks or designated election officials.",
notes:"Polis term-limited. Convention path also available (Dem Mar 28, GOP Apr 11)."},
{s:"CT",n:"Connecticut",filing:"Jun 9, 2026",primary:"Aug 11, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:false,govNote:"Lamont (D)",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Attorney General","Secretary of the State","Treasurer","Comptroller","Full Legislature"],
centralDb:{name:"CT Secretary of the State",url:"https://portal.ct.gov/SOTS/Election-Services/Candidate-and-Committee-Information",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"Town Clerk",localCount:169,localDir:"https://portal.ct.gov/SOTS/Election-Services/Town-Clerks",
counties:[["Fairfield","https://portal.ct.gov/SOTS/Election-Services/Town-Clerks"],["Hartford","https://portal.ct.gov/SOTS/Election-Services/Town-Clerks"],["Litchfield","https://portal.ct.gov/SOTS/Election-Services/Town-Clerks"],["Middlesex","https://portal.ct.gov/SOTS/Election-Services/Town-Clerks"],["New Haven","https://portal.ct.gov/SOTS/Election-Services/Town-Clerks"],["New London","https://portal.ct.gov/SOTS/Election-Services/Town-Clerks"],["Tolland","https://portal.ct.gov/SOTS/Election-Services/Town-Clerks"],["Windham","https://portal.ct.gov/SOTS/Election-Services/Town-Clerks"]],
gap:"All municipal and school board data requires contacting 169 town clerks.",
schoolNote:"Data held by town clerks. Partisan school boards in some districts.",
notes:"No county government. Convention system for nominations."},
{s:"DE",n:"Delaware",filing:"Jul 14, 2026",primary:"Sep 15, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:false,govOpen:false,govNote:"",senate:true,senOpen:false,senNote:"Coons (D)",house:true,
otherOffices:["Attorney General","Treasurer","Legislature"],
centralDb:{name:"DE Election Commissioner",url:"https://elections.delaware.gov/candidate.shtml",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Dept. of Elections",localCount:3,localDir:"https://elections.delaware.gov/",
counties:[["Kent","https://elections.delaware.gov/"],["New Castle","https://elections.delaware.gov/"],["Sussex","https://elections.delaware.gov/"]],
gap:"Only 3 counties — trivial to cover.",
schoolNote:"Data held by individual school district offices.",
notes:"Latest filing deadline in the country (Jul 14). Only 3 counties."},
{s:"FL",n:"Florida",filing:"Apr 24 / Jun 12, 2026",primary:"Aug 18, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:true,govNote:"DeSantis term-limited",senate:true,senOpen:false,senNote:"Special — Rubio resigned; Moody appointed",house:true,
otherOffices:["Attorney General","Chief Financial Officer","Agriculture Commissioner","Legislature"],
centralDb:{name:"FL Candidate Tracking System",url:"https://dos.elections.myflorida.com/candidates/",covers:"Federal, statewide, legislature, multicounty",format:"Searchable database (1996–present)",bulk:false},
localType:"County Supervisor of Elections",localCount:67,localDir:"https://dos.elections.myflorida.com/supervisors/",
counties:[["Alachua","https://www.votealachua.com/"],["Baker","https://www.bakervotes.com/"],["Bay","https://www.bayvotes.org/"],["Bradford","https://www.bradfordelections.com/"],["Brevard","https://www.votebrevard.gov/"],["Broward","https://www.browardvotes.gov/"],["Calhoun","https://dos.elections.myflorida.com/supervisors/"],["Charlotte","https://www.charlottevotes.com/"],["Citrus","https://www.votecitrus.com/"],["Clay","https://www.clayelections.com/"],["Collier","https://www.colliervotes.com/"],["Columbia","https://www.votecolumbia.com/"],["DeSoto","https://www.votedecide.com/"],["Dixie","https://dos.elections.myflorida.com/supervisors/"],["Duval","https://www.duvalelections.gov/"],["Escambia","https://www.escambiavotes.com/"],["Flagler","https://www.flaglerelections.com/"],["Franklin","https://dos.elections.myflorida.com/supervisors/"],["Gadsden","https://dos.elections.myflorida.com/supervisors/"],["Gilchrist","https://dos.elections.myflorida.com/supervisors/"],["Glades","https://dos.elections.myflorida.com/supervisors/"],["Gulf","https://dos.elections.myflorida.com/supervisors/"],["Hamilton","https://dos.elections.myflorida.com/supervisors/"],["Hardee","https://dos.elections.myflorida.com/supervisors/"],["Hendry","https://dos.elections.myflorida.com/supervisors/"],["Hernando","https://www.hernandovotes.gov/"],["Highlands","https://www.highlandselections.com/"],["Hillsborough","https://www.votehillsborough.gov/"],["Holmes","https://dos.elections.myflorida.com/supervisors/"],["Indian River","https://www.voteindianriver.com/"],["Jackson","https://dos.elections.myflorida.com/supervisors/"],["Jefferson","https://dos.elections.myflorida.com/supervisors/"],["Lafayette","https://dos.elections.myflorida.com/supervisors/"],["Lake","https://www.lakevotes.com/"],["Lee","https://www.lee.vote/"],["Leon","https://www.leonvotes.gov/"],["Levy","https://dos.elections.myflorida.com/supervisors/"],["Liberty","https://dos.elections.myflorida.com/supervisors/"],["Madison","https://dos.elections.myflorida.com/supervisors/"],["Manatee","https://www.votemanatee.com/"],["Marion","https://www.votemarion.gov/"],["Martin","https://www.martinvotes.com/"],["Miami-Dade","https://www.miamidadeelections.gov/"],["Monroe","https://www.keyselections.org/"],["Nassau","https://www.nassauvotes.com/"],["Okaloosa","https://www.govoteokaloosa.com/"],["Okeechobee","https://www.voteokeechobee.com/"],["Orange","https://www.ocfelections.com/"],["Osceola","https://www.voteosceola.com/"],["Palm Beach","https://www.pbcelections.org/"],["Pasco","https://www.pascovotes.com/"],["Pinellas","https://www.votepinellas.gov/"],["Polk","https://www.polkelections.gov/"],["Putnam","https://www.putnamsoe.com/"],["Santa Rosa","https://www.santarosavotes.com/"],["Sarasota","https://www.sarasotavotes.gov/"],["Seminole","https://www.voteseminole.gov/"],["St. Johns","https://www.votesjc.gov/"],["St. Lucie","https://www.slcelections.com/"],["Sumter","https://www.sumterelections.org/"],["Suwannee","https://dos.elections.myflorida.com/supervisors/"],["Taylor","https://dos.elections.myflorida.com/supervisors/"],["Union","https://dos.elections.myflorida.com/supervisors/"],["Volusia","https://www.volusia.org/services/government/supervisorofelections/"],["Wakulla","https://dos.elections.myflorida.com/supervisors/"],["Walton","https://www.votewalton.com/"],["Washington","https://dos.elections.myflorida.com/supervisors/"]],
gap:"State system says 'contact your local Supervisor of Elections' for county, municipal, school board races.",
schoolNote:"Data held by county SOE. Nonpartisan elections.",
notes:"DeSantis term-limited. Senate special for Rubio's vacated seat."},
{s:"GA",n:"Georgia",filing:"Mar 6, 2026",primary:"May 19, 2026",runoff:"Jun 16, 2026",general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:true,govNote:"Kemp term-limited",senate:true,senOpen:false,senNote:"Ossoff (D)",house:true,
otherOffices:["Attorney General","Secretary of State","Full Legislature"],
centralDb:{name:"GA Qualified Candidates",url:"https://sos.ga.gov/page/georgia-qualified-candidates",covers:"Federal, statewide, legislature, county (partial)",format:"Downloadable lists",bulk:true},
localType:"County Board of Elections",localCount:159,localDir:"https://elections.sos.ga.gov/Elections/countyelectionoffices.do",
counties:[["Appling","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Atkinson","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Bacon","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Baker","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Baldwin","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Banks","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Barrow","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Bartow","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Ben Hill","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Berrien","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Bibb","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Bleckley","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Brantley","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Brooks","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Bryan","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Bulloch","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Burke","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Butts","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Calhoun","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Camden","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Candler","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Carroll","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Catoosa","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Charlton","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Chatham","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Chattahoochee","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Chattooga","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Cherokee","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Clarke","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Clay","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Clayton","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Clinch","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Cobb","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Coffee","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Colquitt","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Columbia","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Cook","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Coweta","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Crawford","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Crisp","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Dade","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Dawson","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Decatur","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["DeKalb","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Dodge","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Dooly","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Dougherty","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Douglas","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Early","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Echols","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Effingham","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Elbert","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Emanuel","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Evans","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Fannin","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Fayette","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Floyd","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Forsyth","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Franklin","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Fulton","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Gilmer","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Glascock","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Glynn","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Gordon","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Grady","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Greene","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Gwinnett","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Habersham","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Hall","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Hancock","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Haralson","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Harris","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Hart","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Heard","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Henry","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Houston","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Irwin","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Jackson","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Jasper","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Jeff Davis","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Jefferson","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Jenkins","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Johnson","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Jones","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Lamar","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Lanier","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Laurens","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Lee","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Liberty","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Lincoln","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Long","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Lowndes","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Lumpkin","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Macon","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Madison","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Marion","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["McDuffie","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["McIntosh","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Meriwether","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Miller","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Mitchell","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Monroe","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Montgomery","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Morgan","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Murray","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Muscogee","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Newton","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Oconee","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Oglethorpe","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Paulding","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Peach","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Pickens","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Pierce","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Pike","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Polk","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Pulaski","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Putnam","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Quitman","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Rabun","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Randolph","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Richmond","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Rockdale","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Schley","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Screven","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Seminole","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Spalding","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Stephens","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Stewart","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Sumter","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Talbot","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Taliaferro","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Tattnall","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Taylor","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Telfair","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Terrell","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Thomas","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Tift","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Toombs","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Towns","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Treutlen","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Troup","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Turner","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Twiggs","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Union","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Upson","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Walker","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Walton","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Ware","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Warren","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Washington","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Wayne","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Webster","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Wheeler","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["White","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Whitfield","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Wilcox","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Wilkes","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Wilkinson","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"],["Worth","https://elections.sos.ga.gov/Elections/countyelectionoffices.do"]],
gap:"Municipal candidates explicitly excluded. County data only when counties enter it.",
schoolNote:"Data held by county boards of elections.",
notes:"Kemp term-limited. Major battleground. 159 counties (2nd most in US)."},
{s:"HI",n:"Hawaii",filing:"Jun 2, 2026",primary:"Aug 8, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:false,govNote:"Green (D)",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Legislature"],
centralDb:{name:"HI Office of Elections",url:"https://elections.hawaii.gov/candidates/",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Clerk",localCount:4,localDir:"https://elections.hawaii.gov/",
counties:[["Hawaii","https://www.hawaiicounty.gov/departments/elections"],["Honolulu","https://www.honoluluelections.us/"],["Kauai","https://www.kauai.gov/Elections"],["Maui","https://www.mauicounty.gov/239/Office-of-the-County-Clerk"]],
gap:"Essentially none. 4 counties + statewide school board = trivial.",
schoolNote:"★ Statewide elected school board — only state with this. Data at Office of Elections.",
notes:"Easiest state for complete coverage. Only 4 counties."},
{s:"ID",n:"Idaho",filing:"Feb 27, 2026",primary:"May 19, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:false,govNote:"Little (R)",senate:true,senOpen:false,senNote:"Risch (R)",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Supt. of Public Instruction","Full Legislature"],
centralDb:{name:"ID Secretary of State",url:"https://sos.idaho.gov/elections-division/candidate-information/",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Clerk",localCount:44,localDir:"https://sos.idaho.gov/elections-division/",
counties:[["Ada","https://adacounty.id.gov/elections/"],["Adams","https://sos.idaho.gov/elections-division/"],["Bannock","https://www.bannockclerk.com/elections/"],["Bear Lake","https://sos.idaho.gov/elections-division/"],["Benewah","https://sos.idaho.gov/elections-division/"],["Bingham","https://sos.idaho.gov/elections-division/"],["Blaine","https://www.co.blaine.id.us/179/Elections"],["Boise","https://sos.idaho.gov/elections-division/"],["Bonner","https://www.bonnercountyid.gov/elections/"],["Bonneville","https://www.co.bonneville.id.us/elections/"],["Boundary","https://sos.idaho.gov/elections-division/"],["Butte","https://sos.idaho.gov/elections-division/"],["Camas","https://sos.idaho.gov/elections-division/"],["Canyon","https://www.canyoncounty.id.gov/elected-officials/elections"],["Caribou","https://sos.idaho.gov/elections-division/"],["Cassia","https://sos.idaho.gov/elections-division/"],["Clark","https://sos.idaho.gov/elections-division/"],["Clearwater","https://sos.idaho.gov/elections-division/"],["Custer","https://sos.idaho.gov/elections-division/"],["Elmore","https://sos.idaho.gov/elections-division/"],["Franklin","https://sos.idaho.gov/elections-division/"],["Fremont","https://sos.idaho.gov/elections-division/"],["Gem","https://sos.idaho.gov/elections-division/"],["Gooding","https://sos.idaho.gov/elections-division/"],["Idaho","https://sos.idaho.gov/elections-division/"],["Jefferson","https://sos.idaho.gov/elections-division/"],["Jerome","https://sos.idaho.gov/elections-division/"],["Kootenai","https://www.kcgov.us/191/Elections"],["Latah","https://sos.idaho.gov/elections-division/"],["Lemhi","https://sos.idaho.gov/elections-division/"],["Lewis","https://sos.idaho.gov/elections-division/"],["Lincoln","https://sos.idaho.gov/elections-division/"],["Madison","https://sos.idaho.gov/elections-division/"],["Minidoka","https://sos.idaho.gov/elections-division/"],["Nez Perce","https://sos.idaho.gov/elections-division/"],["Oneida","https://sos.idaho.gov/elections-division/"],["Owyhee","https://sos.idaho.gov/elections-division/"],["Payette","https://sos.idaho.gov/elections-division/"],["Power","https://sos.idaho.gov/elections-division/"],["Shoshone","https://sos.idaho.gov/elections-division/"],["Teton","https://sos.idaho.gov/elections-division/"],["Twin Falls","https://www.twinfallscounty.org/Elections"],["Valley","https://sos.idaho.gov/elections-division/"],["Washington","https://sos.idaho.gov/elections-division/"]],
gap:"County, municipal, school board data at county clerk level.",
schoolNote:"Data held by county clerks.",
notes:"Full statewide slate on ballot."},
{s:"IL",n:"Illinois",filing:"Nov 3, 2025",primary:"Mar 17, 2026",runoff:null,general:"Nov 3, 2026",status:"voted",
gov:true,govOpen:false,govNote:"Pritzker (D) — Bailey rematch",senate:true,senOpen:true,senNote:"Open — Durbin retiring; Stratton won Dem nod",house:true,
otherOffices:["Attorney General","Secretary of State","Comptroller","Treasurer","Legislature"],
centralDb:{name:"IL SBE Candidate Search",url:"https://www.elections.il.gov/CandidateSearch.aspx",covers:"★ ALL levels: federal, state, county, municipal, school board, park, library, fire, township, community college",format:"Searchable database",bulk:true},
localType:"County Election Authority (supplemental)",localCount:102,localDir:"https://www.elections.il.gov/ElectionAuthorities.aspx",
counties:[["Adams","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Alexander","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Bond","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Boone","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Brown","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Bureau","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Calhoun","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Carroll","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Cass","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Champaign","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Chicago","https://chicagoelections.gov/"],["Christian","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Clark","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Clay","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Clinton","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Coles","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Cook","https://www.cookcountyclerkil.gov/elections"],["Crawford","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Cumberland","https://www.elections.il.gov/ElectionAuthorities.aspx"],["DeKalb","https://www.elections.il.gov/ElectionAuthorities.aspx"],["De Witt","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Douglas","https://www.elections.il.gov/ElectionAuthorities.aspx"],["DuPage","https://www.dupagecounty.gov/election/"],["Edgar","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Edwards","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Effingham","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Fayette","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Ford","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Franklin","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Fulton","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Gallatin","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Greene","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Grundy","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Hamilton","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Hancock","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Hardin","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Henderson","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Henry","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Iroquois","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Jackson","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Jasper","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Jefferson","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Jersey","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Jo Daviess","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Johnson","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Kane","https://www.kanecountyelections.org/"],["Kankakee","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Kendall","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Knox","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Lake","https://www.lakecountyil.gov/162/County-Clerk"],["La Salle","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Lawrence","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Lee","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Livingston","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Logan","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Macon","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Macoupin","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Madison","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Marion","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Marshall","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Mason","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Massac","https://www.elections.il.gov/ElectionAuthorities.aspx"],["McDonough","https://www.elections.il.gov/ElectionAuthorities.aspx"],["McHenry","https://www.elections.il.gov/ElectionAuthorities.aspx"],["McLean","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Menard","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Mercer","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Monroe","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Montgomery","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Morgan","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Moultrie","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Ogle","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Peoria","https://www.peoriacounty.gov/178/Elections"],["Perry","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Piatt","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Pike","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Pope","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Pulaski","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Putnam","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Randolph","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Richland","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Rock Island","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Saline","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Sangamon","https://www.sangamoncountyclerk.com/elections/"],["Schuyler","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Scott","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Shelby","https://www.elections.il.gov/ElectionAuthorities.aspx"],["St. Clair","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Stark","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Stephenson","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Tazewell","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Union","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Vermilion","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Wabash","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Warren","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Washington","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Wayne","https://www.elections.il.gov/ElectionAuthorities.aspx"],["White","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Whiteside","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Will","https://www.willcountyclerk.com/elections/"],["Williamson","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Winnebago","https://www.elections.il.gov/ElectionAuthorities.aspx"],["Woodford","https://www.elections.il.gov/ElectionAuthorities.aspx"]],
gap:"Minimal — SBE covers virtually all local races.",
schoolNote:"★ School board candidates ARE in the SBE database.",
notes:"★ BEST STATE for local data. Primary held Mar 17. Stratton won Dem Senate nod."},
{s:"IN",n:"Indiana",filing:"Feb 6, 2026",primary:"May 5, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:false,govOpen:false,govNote:"",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Secretary of State","Treasurer","Full Legislature"],
centralDb:{name:"Indiana Voters Portal",url:"https://indianavoters.in.gov/",covers:"Federal, statewide, legislature + ballot preview",format:"Address-based ballot preview",bulk:true},
localType:"County Clerk / Election Board",localCount:92,localDir:"https://www.in.gov/sos/elections/",
counties:[["Adams","https://www.in.gov/sos/elections/"],["Allen","https://www.allencountyvoter.us/"],["Bartholomew","https://www.in.gov/sos/elections/"],["Benton","https://www.in.gov/sos/elections/"],["Blackford","https://www.in.gov/sos/elections/"],["Boone","https://www.in.gov/sos/elections/"],["Brown","https://www.in.gov/sos/elections/"],["Carroll","https://www.in.gov/sos/elections/"],["Cass","https://www.in.gov/sos/elections/"],["Clark","https://www.in.gov/sos/elections/"],["Clay","https://www.in.gov/sos/elections/"],["Clinton","https://www.in.gov/sos/elections/"],["Crawford","https://www.in.gov/sos/elections/"],["Daviess","https://www.in.gov/sos/elections/"],["Dearborn","https://www.in.gov/sos/elections/"],["Decatur","https://www.in.gov/sos/elections/"],["DeKalb","https://www.in.gov/sos/elections/"],["Delaware","https://www.in.gov/sos/elections/"],["Dubois","https://www.in.gov/sos/elections/"],["Elkhart","https://www.in.gov/sos/elections/"],["Fayette","https://www.in.gov/sos/elections/"],["Floyd","https://www.in.gov/sos/elections/"],["Fountain","https://www.in.gov/sos/elections/"],["Franklin","https://www.in.gov/sos/elections/"],["Fulton","https://www.in.gov/sos/elections/"],["Gibson","https://www.in.gov/sos/elections/"],["Grant","https://www.in.gov/sos/elections/"],["Greene","https://www.in.gov/sos/elections/"],["Hamilton","https://www.hamiltoncounty.in.gov/216/Elections"],["Hancock","https://www.in.gov/sos/elections/"],["Harrison","https://www.in.gov/sos/elections/"],["Hendricks","https://www.in.gov/sos/elections/"],["Henry","https://www.in.gov/sos/elections/"],["Howard","https://www.in.gov/sos/elections/"],["Huntington","https://www.in.gov/sos/elections/"],["Jackson","https://www.in.gov/sos/elections/"],["Jasper","https://www.in.gov/sos/elections/"],["Jay","https://www.in.gov/sos/elections/"],["Jefferson","https://www.in.gov/sos/elections/"],["Jennings","https://www.in.gov/sos/elections/"],["Johnson","https://www.in.gov/sos/elections/"],["Knox","https://www.in.gov/sos/elections/"],["Kosciusko","https://www.in.gov/sos/elections/"],["LaGrange","https://www.in.gov/sos/elections/"],["Lake","https://www.lakecountyin.org/departments/voter_registration/"],["LaPorte","https://www.in.gov/sos/elections/"],["Lawrence","https://www.in.gov/sos/elections/"],["Madison","https://www.in.gov/sos/elections/"],["Marion","https://vote.indy.gov/"],["Marshall","https://www.in.gov/sos/elections/"],["Martin","https://www.in.gov/sos/elections/"],["Miami","https://www.in.gov/sos/elections/"],["Monroe","https://www.in.gov/sos/elections/"],["Montgomery","https://www.in.gov/sos/elections/"],["Morgan","https://www.in.gov/sos/elections/"],["Newton","https://www.in.gov/sos/elections/"],["Noble","https://www.in.gov/sos/elections/"],["Ohio","https://www.in.gov/sos/elections/"],["Orange","https://www.in.gov/sos/elections/"],["Owen","https://www.in.gov/sos/elections/"],["Parke","https://www.in.gov/sos/elections/"],["Perry","https://www.in.gov/sos/elections/"],["Pike","https://www.in.gov/sos/elections/"],["Porter","https://www.in.gov/sos/elections/"],["Posey","https://www.in.gov/sos/elections/"],["Pulaski","https://www.in.gov/sos/elections/"],["Putnam","https://www.in.gov/sos/elections/"],["Randolph","https://www.in.gov/sos/elections/"],["Ripley","https://www.in.gov/sos/elections/"],["Rush","https://www.in.gov/sos/elections/"],["Scott","https://www.in.gov/sos/elections/"],["Shelby","https://www.in.gov/sos/elections/"],["Spencer","https://www.in.gov/sos/elections/"],["St. Joseph","https://www.sjcindiana.com/508/Elections"],["Starke","https://www.in.gov/sos/elections/"],["Steuben","https://www.in.gov/sos/elections/"],["Sullivan","https://www.in.gov/sos/elections/"],["Switzerland","https://www.in.gov/sos/elections/"],["Tippecanoe","https://www.tippecanoe.in.gov/1059/Election-Board"],["Tipton","https://www.in.gov/sos/elections/"],["Union","https://www.in.gov/sos/elections/"],["Vanderburgh","https://www.evansvillegov.org/government/department-index/county-clerk-s-office/voter-registration/"],["Vermillion","https://www.in.gov/sos/elections/"],["Vigo","https://www.in.gov/sos/elections/"],["Wabash","https://www.in.gov/sos/elections/"],["Warren","https://www.in.gov/sos/elections/"],["Warrick","https://www.in.gov/sos/elections/"],["Washington","https://www.in.gov/sos/elections/"],["Wayne","https://www.in.gov/sos/elections/"],["Wells","https://www.in.gov/sos/elections/"],["White","https://www.in.gov/sos/elections/"],["Whitley","https://www.in.gov/sos/elections/"]],
gap:"County, municipal, school board at county level.",
schoolNote:"Some elected, some appointed. Elected ones file with county board.",
notes:"No Governor or Senate race. House, statewide offices, legislature."},
{s:"IA",n:"Iowa",filing:"Mar 13, 2026",primary:"Jun 2, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:true,govNote:"Reynolds not running",senate:true,senOpen:true,senNote:"Open — Ernst retiring",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Auditor","Sec. of Agriculture","Legislature"],
centralDb:{name:"IA SoS Candidate Search",url:"https://sos.iowa.gov/elections/candidates/search.html",covers:"Federal, statewide, legislature",format:"Searchable database",bulk:false},
localType:"County Auditor",localCount:99,localDir:"https://sos.iowa.gov/elections/auditors/auditorslist.html",
counties:[["Adair","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Adams","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Allamakee","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Appanoose","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Audubon","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Benton","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Black Hawk","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Boone","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Bremer","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Buchanan","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Buena Vista","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Butler","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Calhoun","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Carroll","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Cass","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Cedar","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Cerro Gordo","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Cherokee","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Chickasaw","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Clarke","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Clay","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Clayton","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Clinton","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Crawford","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Dallas","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Davis","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Decatur","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Delaware","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Des Moines","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Dickinson","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Dubuque","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Emmet","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Fayette","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Floyd","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Franklin","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Fremont","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Greene","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Grundy","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Guthrie","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Hamilton","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Hancock","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Hardin","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Harrison","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Henry","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Howard","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Humboldt","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Ida","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Iowa","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Jackson","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Jasper","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Jefferson","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Johnson","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Jones","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Keokuk","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Kossuth","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Lee","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Linn","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Louisa","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Lucas","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Lyon","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Madison","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Mahaska","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Marion","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Marshall","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Mills","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Mitchell","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Monona","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Monroe","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Montgomery","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Muscatine","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["O'Brien","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Osceola","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Page","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Palo Alto","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Plymouth","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Pocahontas","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Polk","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Pottawattamie","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Poweshiek","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Ringgold","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Sac","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Scott","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Shelby","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Sioux","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Story","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Tama","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Taylor","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Union","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Van Buren","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Wapello","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Warren","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Washington","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Wayne","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Webster","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Winnebago","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Winneshiek","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Woodbury","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Worth","https://sos.iowa.gov/elections/auditors/auditorslist.html"],["Wright","https://sos.iowa.gov/elections/auditors/auditorslist.html"]],
gap:"County, municipal, school board data at county auditor level.",
schoolNote:"School board elections typically September (separate). Data at county auditors.",
notes:"Both Gov and Senate open. Convention backstop if no candidate exceeds 35%."},
{s:"KS",n:"Kansas",filing:"Jun 1, 2026",primary:"Aug 4, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:true,govNote:"Kelly (D) term-limited",senate:true,senOpen:false,senNote:"Marshall (R)",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Insurance Commissioner","House-only Legislature"],
centralDb:{name:"KS Secretary of State",url:"https://sos.ks.gov/elections/elections-candidates.html",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Election Officer",localCount:105,localDir:"https://sos.ks.gov/elections/county-election-officers.html",
counties:[["Allen","https://sos.ks.gov/elections/county-election-officers.html"],["Anderson","https://sos.ks.gov/elections/county-election-officers.html"],["Atchison","https://sos.ks.gov/elections/county-election-officers.html"],["Barber","https://sos.ks.gov/elections/county-election-officers.html"],["Barton","https://sos.ks.gov/elections/county-election-officers.html"],["Bourbon","https://sos.ks.gov/elections/county-election-officers.html"],["Brown","https://sos.ks.gov/elections/county-election-officers.html"],["Butler","https://sos.ks.gov/elections/county-election-officers.html"],["Chase","https://sos.ks.gov/elections/county-election-officers.html"],["Chautauqua","https://sos.ks.gov/elections/county-election-officers.html"],["Cherokee","https://sos.ks.gov/elections/county-election-officers.html"],["Cheyenne","https://sos.ks.gov/elections/county-election-officers.html"],["Clark","https://sos.ks.gov/elections/county-election-officers.html"],["Clay","https://sos.ks.gov/elections/county-election-officers.html"],["Cloud","https://sos.ks.gov/elections/county-election-officers.html"],["Coffey","https://sos.ks.gov/elections/county-election-officers.html"],["Comanche","https://sos.ks.gov/elections/county-election-officers.html"],["Cowley","https://sos.ks.gov/elections/county-election-officers.html"],["Crawford","https://sos.ks.gov/elections/county-election-officers.html"],["Decatur","https://sos.ks.gov/elections/county-election-officers.html"],["Dickinson","https://sos.ks.gov/elections/county-election-officers.html"],["Doniphan","https://sos.ks.gov/elections/county-election-officers.html"],["Douglas","https://sos.ks.gov/elections/county-election-officers.html"],["Edwards","https://sos.ks.gov/elections/county-election-officers.html"],["Elk","https://sos.ks.gov/elections/county-election-officers.html"],["Ellis","https://sos.ks.gov/elections/county-election-officers.html"],["Ellsworth","https://sos.ks.gov/elections/county-election-officers.html"],["Finney","https://sos.ks.gov/elections/county-election-officers.html"],["Ford","https://sos.ks.gov/elections/county-election-officers.html"],["Franklin","https://sos.ks.gov/elections/county-election-officers.html"],["Geary","https://sos.ks.gov/elections/county-election-officers.html"],["Gove","https://sos.ks.gov/elections/county-election-officers.html"],["Graham","https://sos.ks.gov/elections/county-election-officers.html"],["Grant","https://sos.ks.gov/elections/county-election-officers.html"],["Gray","https://sos.ks.gov/elections/county-election-officers.html"],["Greeley","https://sos.ks.gov/elections/county-election-officers.html"],["Greenwood","https://sos.ks.gov/elections/county-election-officers.html"],["Hamilton","https://sos.ks.gov/elections/county-election-officers.html"],["Harper","https://sos.ks.gov/elections/county-election-officers.html"],["Harvey","https://sos.ks.gov/elections/county-election-officers.html"],["Haskell","https://sos.ks.gov/elections/county-election-officers.html"],["Hodgeman","https://sos.ks.gov/elections/county-election-officers.html"],["Jackson","https://sos.ks.gov/elections/county-election-officers.html"],["Jefferson","https://sos.ks.gov/elections/county-election-officers.html"],["Jewell","https://sos.ks.gov/elections/county-election-officers.html"],["Johnson","https://sos.ks.gov/elections/county-election-officers.html"],["Kearny","https://sos.ks.gov/elections/county-election-officers.html"],["Kingman","https://sos.ks.gov/elections/county-election-officers.html"],["Kiowa","https://sos.ks.gov/elections/county-election-officers.html"],["Labette","https://sos.ks.gov/elections/county-election-officers.html"],["Lane","https://sos.ks.gov/elections/county-election-officers.html"],["Leavenworth","https://sos.ks.gov/elections/county-election-officers.html"],["Lincoln","https://sos.ks.gov/elections/county-election-officers.html"],["Linn","https://sos.ks.gov/elections/county-election-officers.html"],["Logan","https://sos.ks.gov/elections/county-election-officers.html"],["Lyon","https://sos.ks.gov/elections/county-election-officers.html"],["Marion","https://sos.ks.gov/elections/county-election-officers.html"],["Marshall","https://sos.ks.gov/elections/county-election-officers.html"],["McPherson","https://sos.ks.gov/elections/county-election-officers.html"],["Meade","https://sos.ks.gov/elections/county-election-officers.html"],["Miami","https://sos.ks.gov/elections/county-election-officers.html"],["Mitchell","https://sos.ks.gov/elections/county-election-officers.html"],["Montgomery","https://sos.ks.gov/elections/county-election-officers.html"],["Morris","https://sos.ks.gov/elections/county-election-officers.html"],["Morton","https://sos.ks.gov/elections/county-election-officers.html"],["Nemaha","https://sos.ks.gov/elections/county-election-officers.html"],["Neosho","https://sos.ks.gov/elections/county-election-officers.html"],["Ness","https://sos.ks.gov/elections/county-election-officers.html"],["Norton","https://sos.ks.gov/elections/county-election-officers.html"],["Osage","https://sos.ks.gov/elections/county-election-officers.html"],["Osborne","https://sos.ks.gov/elections/county-election-officers.html"],["Ottawa","https://sos.ks.gov/elections/county-election-officers.html"],["Pawnee","https://sos.ks.gov/elections/county-election-officers.html"],["Phillips","https://sos.ks.gov/elections/county-election-officers.html"],["Pottawatomie","https://sos.ks.gov/elections/county-election-officers.html"],["Pratt","https://sos.ks.gov/elections/county-election-officers.html"],["Rawlins","https://sos.ks.gov/elections/county-election-officers.html"],["Reno","https://sos.ks.gov/elections/county-election-officers.html"],["Republic","https://sos.ks.gov/elections/county-election-officers.html"],["Rice","https://sos.ks.gov/elections/county-election-officers.html"],["Riley","https://sos.ks.gov/elections/county-election-officers.html"],["Rooks","https://sos.ks.gov/elections/county-election-officers.html"],["Rush","https://sos.ks.gov/elections/county-election-officers.html"],["Russell","https://sos.ks.gov/elections/county-election-officers.html"],["Saline","https://sos.ks.gov/elections/county-election-officers.html"],["Scott","https://sos.ks.gov/elections/county-election-officers.html"],["Sedgwick","https://sos.ks.gov/elections/county-election-officers.html"],["Seward","https://sos.ks.gov/elections/county-election-officers.html"],["Shawnee","https://sos.ks.gov/elections/county-election-officers.html"],["Sheridan","https://sos.ks.gov/elections/county-election-officers.html"],["Sherman","https://sos.ks.gov/elections/county-election-officers.html"],["Smith","https://sos.ks.gov/elections/county-election-officers.html"],["Stafford","https://sos.ks.gov/elections/county-election-officers.html"],["Stanton","https://sos.ks.gov/elections/county-election-officers.html"],["Stevens","https://sos.ks.gov/elections/county-election-officers.html"],["Sumner","https://sos.ks.gov/elections/county-election-officers.html"],["Thomas","https://sos.ks.gov/elections/county-election-officers.html"],["Trego","https://sos.ks.gov/elections/county-election-officers.html"],["Wabaunsee","https://sos.ks.gov/elections/county-election-officers.html"],["Wallace","https://sos.ks.gov/elections/county-election-officers.html"],["Washington","https://sos.ks.gov/elections/county-election-officers.html"],["Wichita","https://sos.ks.gov/elections/county-election-officers.html"],["Wilson","https://sos.ks.gov/elections/county-election-officers.html"],["Woodson","https://sos.ks.gov/elections/county-election-officers.html"],["Wyandotte","https://sos.ks.gov/elections/county-election-officers.html"]],
gap:"County, municipal, school board data at county election officer level.",
schoolNote:"Data held by county election officers.",
notes:"Kelly (D) term-limited. Open governor's race in red-leaning state."},
{s:"KY",n:"Kentucky",filing:"Jan 9, 2026",primary:"May 19, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:false,govOpen:false,govNote:"",senate:true,senOpen:true,senNote:"Open — McConnell retiring",house:true,
otherOffices:["Legislature"],
centralDb:{name:"KY State Board of Elections",url:"https://elect.ky.gov/candidates/Pages/default.aspx",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Clerk",localCount:120,localDir:"https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx",
counties:[["Adair","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Allen","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Anderson","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Ballard","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Barren","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Bath","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Bell","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Boone","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Bourbon","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Boyd","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Boyle","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Bracken","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Breathitt","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Breckinridge","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Bullitt","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Butler","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Caldwell","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Calloway","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Campbell","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Carlisle","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Carroll","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Carter","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Casey","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Christian","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Clark","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Clay","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Clinton","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Crittenden","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Cumberland","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Daviess","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Edmonson","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Elliott","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Estill","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Fayette","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Fleming","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Floyd","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Franklin","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Fulton","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Gallatin","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Garrard","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Grant","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Graves","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Grayson","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Green","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Greenup","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Hancock","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Hardin","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Harlan","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Harrison","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Hart","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Henderson","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Henry","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Hickman","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Hopkins","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Jackson","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Jefferson","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Jessamine","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Johnson","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Kenton","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Knott","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Knox","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Larue","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Laurel","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Lawrence","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Lee","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Leslie","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Letcher","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Lewis","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Lincoln","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Livingston","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Logan","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Lyon","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["McCracken","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["McCreary","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["McLean","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Madison","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Magoffin","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Marion","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Marshall","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Martin","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Mason","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Meade","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Menifee","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Mercer","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Metcalfe","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Monroe","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Montgomery","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Morgan","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Muhlenberg","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Nelson","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Nicholas","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Ohio","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Oldham","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Owen","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Owsley","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Pendleton","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Perry","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Pike","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Powell","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Pulaski","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Robertson","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Rockcastle","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Rowan","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Russell","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Scott","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Shelby","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Simpson","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Spencer","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Taylor","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Todd","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Trigg","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Trimble","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Union","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Warren","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Washington","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Wayne","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Webster","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Whitley","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Wolfe","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"],["Woodford","https://elect.ky.gov/About-Us/Pages/County-Clerks.aspx"]],
gap:"County, municipal, school board data at county clerk level. 120 counties.",
schoolNote:"Data held by county clerks. Nonpartisan races.",
notes:"McConnell retiring after decades as Senate leader. Earliest filing deadline (Jan 9)."},
{s:"LA",n:"Louisiana",filing:"Feb 13 / Jul 31, 2026",primary:"May 16 / Oct 24, 2026",runoff:"Jun 27 / Nov 3",general:"Nov 3, 2026",status:"partial",
gov:false,govOpen:false,govNote:"",senate:true,senOpen:false,senNote:"Cassidy (R)",house:true,
otherOffices:["No legislative races (off-cycle)"],
centralDb:{name:"LA GeauxVote Voter Portal",url:"https://voterportal.sos.la.gov/Graphical",covers:"★ ALL levels: federal, state, parish, municipal, school board, judicial, special district",format:"Address-based full ballot lookup",bulk:true},
localType:"N/A — Fully centralized",localCount:0,localDir:"https://voterportal.sos.la.gov/Graphical",
gap:"NONE — Louisiana centralizes ALL candidate data in one system.",
schoolNote:"★ School board candidates ARE in the state system. Fully centralized.",
notes:"★ GOLD STANDARD for centralized data. New party primary for congressional. Old system for others."},
{s:"ME",n:"Maine",filing:"Mar 16, 2026",primary:"Jun 9, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:true,govNote:"Mills (D) term-limited",senate:true,senOpen:false,senNote:"Collins (R)",house:true,
otherOffices:["Treasurer","Full Legislature"],
centralDb:{name:"ME Secretary of State",url:"https://www.maine.gov/sos/elections-voting/becoming-a-candidate",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"Municipal Clerk",localCount:486,localDir:"https://www.maine.gov/sos/cec/elec/munic.html",
counties:[["Androscoggin","https://www.maine.gov/sos/cec/elec/munic.html"],["Aroostook","https://www.maine.gov/sos/cec/elec/munic.html"],["Cumberland","https://www.maine.gov/sos/cec/elec/munic.html"],["Franklin","https://www.maine.gov/sos/cec/elec/munic.html"],["Hancock","https://www.maine.gov/sos/cec/elec/munic.html"],["Kennebec","https://www.maine.gov/sos/cec/elec/munic.html"],["Knox","https://www.maine.gov/sos/cec/elec/munic.html"],["Lincoln","https://www.maine.gov/sos/cec/elec/munic.html"],["Oxford","https://www.maine.gov/sos/cec/elec/munic.html"],["Penobscot","https://www.maine.gov/sos/cec/elec/munic.html"],["Piscataquis","https://www.maine.gov/sos/cec/elec/munic.html"],["Sagadahoc","https://www.maine.gov/sos/cec/elec/munic.html"],["Somerset","https://www.maine.gov/sos/cec/elec/munic.html"],["Waldo","https://www.maine.gov/sos/cec/elec/munic.html"],["Washington","https://www.maine.gov/sos/cec/elec/munic.html"],["York","https://www.maine.gov/sos/cec/elec/munic.html"]],
gap:"All local data at municipal level. 486 municipalities — labor-intensive.",
schoolNote:"Data held by municipal or school district clerks.",
notes:"Mills (D) term-limited. Ranked-choice voting. Collins (R) faces competitive re-election."},
{s:"MD",n:"Maryland",filing:"Feb 24, 2026",primary:"Jun 23, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:false,govNote:"Moore (D)",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Attorney General","Comptroller","Full Legislature"],
centralDb:{name:"MD State Board of Elections",url:"https://elections.maryland.gov/elections/2026/primary_candidates/index.html",covers:"Federal, statewide, legislature",format:"Online listings + PDF/CSV",bulk:true},
localType:"County Board of Elections",localCount:24,localDir:"https://elections.maryland.gov/about/county_boards.html",
counties:[["Allegany","https://www.alleganygov.org/158/Election-Office"],["Anne Arundel","https://www.aacounty.org/boards-and-commissions/board-of-elections/"],["Baltimore City","https://boe.baltimorecity.gov"],["Baltimore County","https://www.baltimorecountymd.gov/departments/elections"],["Calvert","https://www.calvertcountymd.gov/154/Election-Board"],["Caroline","https://www.carolinemdelections.org/"],["Carroll","https://elections.carrollcountymd.gov/"],["Cecil","https://www.ccgov.org/government/election-board"],["Charles","https://www.charlescountymd.gov/government/board-of-elections"],["Dorchester","https://www.docomdelections.org/"],["Frederick","https://www.frederickcountymd.gov/1198/Board-of-Elections"],["Garrett","https://www.garrettcounty.org/board-of-elections"],["Harford","https://www.harfordvotes.gov"],["Howard","https://www.howardcountymd.gov/boards-commissions/board-elections"],["Kent","https://www.kentcountyelections.org"],["Montgomery","https://www.montgomerycountymd.gov/Elections/"],["Prince George's","https://www.princegeorgescountymd.gov/boards-commissions/board-elections"],["Queen Anne's","https://www.qacelections.com/"],["Somerset","https://www.somersetmd.us/government/board_of_elections.php"],["St. Mary's","https://www.stmarysmd.com/supervisorofelections/"],["Talbot","https://www.talbotcountymd.gov/index.php?page=Election_Board"],["Washington","https://www.washco-mdelections.org/"],["Wicomico","https://www.wicomicocounty.org/132/Board-of-Elections"],["Worcester","https://www.co.worcester.md.us/departments/board-elections"]],
gap:"County, municipal, school board at county level. Only 24 — manageable.",
schoolNote:"Some counties elect, some appoint school boards. Check county board.",
notes:"No Senate race. Moore (D) running for re-election. Only 24 jurisdictions."},
{s:"MA",n:"Massachusetts",filing:"Jun 2, 2026",primary:"Sep 1, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:false,govNote:"Healey (D)",senate:true,senOpen:false,senNote:"Markey (D)",house:true,
otherOffices:["Attorney General","Secretary of the Commonwealth","Treasurer","Auditor","Full Legislature"],
centralDb:{name:"MA Secretary of the Commonwealth",url:"https://www.sec.state.ma.us/divisions/elections/candidates.htm",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"City/Town Clerk",localCount:351,localDir:"https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm",
counties:[["Barnstable","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Berkshire","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Bristol","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Dukes","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Essex","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Franklin","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Hampden","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Hampshire","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Middlesex","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Nantucket","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Norfolk","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Plymouth","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Suffolk","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"],["Worcester","https://www.sec.state.ma.us/divisions/elections/town-city-clerks.htm"]],
gap:"All local and school committee data requires contacting 351 clerks.",
schoolNote:"Called 'school committees.' Data held by city/town clerks.",
notes:"Full slate of statewide offices plus Senate."},
{s:"MI",n:"Michigan",filing:"Apr 21, 2026",primary:"Aug 4, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:true,govNote:"Whitmer term-limited",senate:true,senOpen:true,senNote:"Open — Peters retiring",house:true,
otherOffices:["Attorney General","Secretary of State","Full Legislature"],
centralDb:{name:"MI Voter Information Center (MVIC)",url:"https://mvic.sos.state.mi.us/Candidates",covers:"Federal, statewide, legislature + ballot preview",format:"Address-based ballot preview",bulk:false},
localType:"County Clerk",localCount:83,localDir:"https://mvic.sos.state.mi.us/Clerk",
counties:[["Alcona","https://mvic.sos.state.mi.us/Clerk"],["Alger","https://mvic.sos.state.mi.us/Clerk"],["Allegan","https://mvic.sos.state.mi.us/Clerk"],["Alpena","https://mvic.sos.state.mi.us/Clerk"],["Antrim","https://mvic.sos.state.mi.us/Clerk"],["Arenac","https://mvic.sos.state.mi.us/Clerk"],["Baraga","https://mvic.sos.state.mi.us/Clerk"],["Barry","https://mvic.sos.state.mi.us/Clerk"],["Bay","https://mvic.sos.state.mi.us/Clerk"],["Benzie","https://mvic.sos.state.mi.us/Clerk"],["Berrien","https://mvic.sos.state.mi.us/Clerk"],["Branch","https://mvic.sos.state.mi.us/Clerk"],["Calhoun","https://mvic.sos.state.mi.us/Clerk"],["Cass","https://mvic.sos.state.mi.us/Clerk"],["Charlevoix","https://mvic.sos.state.mi.us/Clerk"],["Cheboygan","https://mvic.sos.state.mi.us/Clerk"],["Chippewa","https://mvic.sos.state.mi.us/Clerk"],["Clare","https://mvic.sos.state.mi.us/Clerk"],["Clinton","https://mvic.sos.state.mi.us/Clerk"],["Crawford","https://mvic.sos.state.mi.us/Clerk"],["Delta","https://mvic.sos.state.mi.us/Clerk"],["Dickinson","https://mvic.sos.state.mi.us/Clerk"],["Eaton","https://mvic.sos.state.mi.us/Clerk"],["Emmet","https://mvic.sos.state.mi.us/Clerk"],["Genesee","https://mvic.sos.state.mi.us/Clerk"],["Gladwin","https://mvic.sos.state.mi.us/Clerk"],["Gogebic","https://mvic.sos.state.mi.us/Clerk"],["Grand Traverse","https://mvic.sos.state.mi.us/Clerk"],["Gratiot","https://mvic.sos.state.mi.us/Clerk"],["Hillsdale","https://mvic.sos.state.mi.us/Clerk"],["Houghton","https://mvic.sos.state.mi.us/Clerk"],["Huron","https://mvic.sos.state.mi.us/Clerk"],["Ingham","https://mvic.sos.state.mi.us/Clerk"],["Ionia","https://mvic.sos.state.mi.us/Clerk"],["Iosco","https://mvic.sos.state.mi.us/Clerk"],["Iron","https://mvic.sos.state.mi.us/Clerk"],["Isabella","https://mvic.sos.state.mi.us/Clerk"],["Jackson","https://mvic.sos.state.mi.us/Clerk"],["Kalamazoo","https://mvic.sos.state.mi.us/Clerk"],["Kalkaska","https://mvic.sos.state.mi.us/Clerk"],["Kent","https://mvic.sos.state.mi.us/Clerk"],["Keweenaw","https://mvic.sos.state.mi.us/Clerk"],["Lake","https://mvic.sos.state.mi.us/Clerk"],["Lapeer","https://mvic.sos.state.mi.us/Clerk"],["Leelanau","https://mvic.sos.state.mi.us/Clerk"],["Lenawee","https://mvic.sos.state.mi.us/Clerk"],["Livingston","https://mvic.sos.state.mi.us/Clerk"],["Luce","https://mvic.sos.state.mi.us/Clerk"],["Mackinac","https://mvic.sos.state.mi.us/Clerk"],["Macomb","https://mvic.sos.state.mi.us/Clerk"],["Manistee","https://mvic.sos.state.mi.us/Clerk"],["Marquette","https://mvic.sos.state.mi.us/Clerk"],["Mason","https://mvic.sos.state.mi.us/Clerk"],["Mecosta","https://mvic.sos.state.mi.us/Clerk"],["Menominee","https://mvic.sos.state.mi.us/Clerk"],["Midland","https://mvic.sos.state.mi.us/Clerk"],["Missaukee","https://mvic.sos.state.mi.us/Clerk"],["Monroe","https://mvic.sos.state.mi.us/Clerk"],["Montcalm","https://mvic.sos.state.mi.us/Clerk"],["Montmorency","https://mvic.sos.state.mi.us/Clerk"],["Muskegon","https://mvic.sos.state.mi.us/Clerk"],["Newaygo","https://mvic.sos.state.mi.us/Clerk"],["Oakland","https://mvic.sos.state.mi.us/Clerk"],["Oceana","https://mvic.sos.state.mi.us/Clerk"],["Ogemaw","https://mvic.sos.state.mi.us/Clerk"],["Ontonagon","https://mvic.sos.state.mi.us/Clerk"],["Osceola","https://mvic.sos.state.mi.us/Clerk"],["Oscoda","https://mvic.sos.state.mi.us/Clerk"],["Otsego","https://mvic.sos.state.mi.us/Clerk"],["Ottawa","https://mvic.sos.state.mi.us/Clerk"],["Presque Isle","https://mvic.sos.state.mi.us/Clerk"],["Roscommon","https://mvic.sos.state.mi.us/Clerk"],["Saginaw","https://mvic.sos.state.mi.us/Clerk"],["Sanilac","https://mvic.sos.state.mi.us/Clerk"],["Schoolcraft","https://mvic.sos.state.mi.us/Clerk"],["Shiawassee","https://mvic.sos.state.mi.us/Clerk"],["St. Clair","https://mvic.sos.state.mi.us/Clerk"],["St. Joseph","https://mvic.sos.state.mi.us/Clerk"],["Tuscola","https://mvic.sos.state.mi.us/Clerk"],["Van Buren","https://mvic.sos.state.mi.us/Clerk"],["Washtenaw","https://mvic.sos.state.mi.us/Clerk"],["Wayne","https://mvic.sos.state.mi.us/Clerk"],["Wexford","https://mvic.sos.state.mi.us/Clerk"]],
gap:"MVIC shows local candidates closer to election. For early data, contact county clerks.",
schoolNote:"Data held by county clerks or school district secretaries.",
notes:"Whitmer term-limited. Peters retiring. Both top-of-ticket races open — major battleground."},
{s:"MN",n:"Minnesota",filing:"Jun 2, 2026",primary:"Aug 11, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:true,govNote:"Walz not running",senate:true,senOpen:true,senNote:"Open — Smith retiring",house:true,
otherOffices:["Attorney General","Secretary of State","Auditor","Full Legislature"],
centralDb:{name:"MN SoS Candidate Filing",url:"https://candidates.sos.mn.gov/",covers:"★ ALL levels: federal, state, county, municipal, school district, judicial, township",format:"Searchable database + downloadable text files",bulk:true},
localType:"County Auditor (supplemental)",localCount:87,localDir:"https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/",
counties:[["Aitkin","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Anoka","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Becker","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Beltrami","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Benton","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Big Stone","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Blue Earth","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Brown","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Carlton","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Carver","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Cass","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Chippewa","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Chisago","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Clay","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Clearwater","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Cook","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Cottonwood","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Crow Wing","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Dakota","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Dodge","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Douglas","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Faribault","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Fillmore","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Freeborn","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Goodhue","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Grant","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Hennepin","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Houston","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Hubbard","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Isanti","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Itasca","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Jackson","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Kanabec","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Kandiyohi","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Kittson","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Koochiching","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Lac qui Parle","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Lake","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Lake of the Woods","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Le Sueur","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Lincoln","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Lyon","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Mahnomen","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Marshall","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Martin","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["McLeod","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Meeker","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Mille Lacs","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Morrison","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Mower","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Murray","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Nicollet","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Nobles","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Norman","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Olmsted","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Otter Tail","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Pennington","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Pine","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Pipestone","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Polk","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Pope","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Ramsey","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Red Lake","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Redwood","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Renville","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Rice","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Rock","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Roseau","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Scott","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Sherburne","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Sibley","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["St. Louis","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Stearns","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Steele","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Stevens","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Swift","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Todd","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Traverse","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Wabasha","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Wadena","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Waseca","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Washington","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Watonwan","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Wilkin","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Winona","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Wright","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"],["Yellow Medicine","https://www.sos.state.mn.us/election-administration-campaigns/county-auditors/"]],
gap:"Minimal — state system covers most local races.",
schoolNote:"★ School district candidates ARE in the state system.",
notes:"★ EXCELLENT centralized data. Both Gov and Senate open. Walz was 2024 VP nominee."},
{s:"MS",n:"Mississippi",filing:"Dec 26, 2025",primary:"Mar 10, 2026",runoff:"Apr 7, 2026",general:"Nov 3, 2026",status:"voted",
gov:false,govOpen:false,govNote:"",senate:true,senOpen:false,senNote:"Hyde-Smith (R) renominated",house:true,
otherOffices:["No legislative races (off-cycle)"],
centralDb:{name:"MS Secretary of State",url:"https://www.sos.ms.gov/elections-voting/candidate-qualifying",covers:"Federal, statewide (when on-cycle)",format:"Online listings",bulk:false},
localType:"County Circuit Clerk",localCount:82,localDir:"https://www.sos.ms.gov/elections-voting",
counties:[["Adams","https://www.sos.ms.gov/elections-voting"],["Alcorn","https://www.sos.ms.gov/elections-voting"],["Amite","https://www.sos.ms.gov/elections-voting"],["Attala","https://www.sos.ms.gov/elections-voting"],["Benton","https://www.sos.ms.gov/elections-voting"],["Bolivar","https://www.sos.ms.gov/elections-voting"],["Calhoun","https://www.sos.ms.gov/elections-voting"],["Carroll","https://www.sos.ms.gov/elections-voting"],["Chickasaw","https://www.sos.ms.gov/elections-voting"],["Choctaw","https://www.sos.ms.gov/elections-voting"],["Claiborne","https://www.sos.ms.gov/elections-voting"],["Clarke","https://www.sos.ms.gov/elections-voting"],["Clay","https://www.sos.ms.gov/elections-voting"],["Coahoma","https://www.sos.ms.gov/elections-voting"],["Copiah","https://www.sos.ms.gov/elections-voting"],["Covington","https://www.sos.ms.gov/elections-voting"],["DeSoto","https://www.sos.ms.gov/elections-voting"],["Forrest","https://www.sos.ms.gov/elections-voting"],["Franklin","https://www.sos.ms.gov/elections-voting"],["George","https://www.sos.ms.gov/elections-voting"],["Greene","https://www.sos.ms.gov/elections-voting"],["Grenada","https://www.sos.ms.gov/elections-voting"],["Hancock","https://www.sos.ms.gov/elections-voting"],["Harrison","https://www.sos.ms.gov/elections-voting"],["Hinds","https://www.sos.ms.gov/elections-voting"],["Holmes","https://www.sos.ms.gov/elections-voting"],["Humphreys","https://www.sos.ms.gov/elections-voting"],["Issaquena","https://www.sos.ms.gov/elections-voting"],["Itawamba","https://www.sos.ms.gov/elections-voting"],["Jackson","https://www.sos.ms.gov/elections-voting"],["Jasper","https://www.sos.ms.gov/elections-voting"],["Jefferson","https://www.sos.ms.gov/elections-voting"],["Jefferson Davis","https://www.sos.ms.gov/elections-voting"],["Jones","https://www.sos.ms.gov/elections-voting"],["Kemper","https://www.sos.ms.gov/elections-voting"],["Lafayette","https://www.sos.ms.gov/elections-voting"],["Lamar","https://www.sos.ms.gov/elections-voting"],["Lauderdale","https://www.sos.ms.gov/elections-voting"],["Lawrence","https://www.sos.ms.gov/elections-voting"],["Leake","https://www.sos.ms.gov/elections-voting"],["Lee","https://www.sos.ms.gov/elections-voting"],["Leflore","https://www.sos.ms.gov/elections-voting"],["Lincoln","https://www.sos.ms.gov/elections-voting"],["Lowndes","https://www.sos.ms.gov/elections-voting"],["Madison","https://www.sos.ms.gov/elections-voting"],["Marion","https://www.sos.ms.gov/elections-voting"],["Marshall","https://www.sos.ms.gov/elections-voting"],["Monroe","https://www.sos.ms.gov/elections-voting"],["Montgomery","https://www.sos.ms.gov/elections-voting"],["Neshoba","https://www.sos.ms.gov/elections-voting"],["Newton","https://www.sos.ms.gov/elections-voting"],["Noxubee","https://www.sos.ms.gov/elections-voting"],["Oktibbeha","https://www.sos.ms.gov/elections-voting"],["Panola","https://www.sos.ms.gov/elections-voting"],["Pearl River","https://www.sos.ms.gov/elections-voting"],["Perry","https://www.sos.ms.gov/elections-voting"],["Pike","https://www.sos.ms.gov/elections-voting"],["Pontotoc","https://www.sos.ms.gov/elections-voting"],["Prentiss","https://www.sos.ms.gov/elections-voting"],["Quitman","https://www.sos.ms.gov/elections-voting"],["Rankin","https://www.sos.ms.gov/elections-voting"],["Scott","https://www.sos.ms.gov/elections-voting"],["Sharkey","https://www.sos.ms.gov/elections-voting"],["Simpson","https://www.sos.ms.gov/elections-voting"],["Smith","https://www.sos.ms.gov/elections-voting"],["Stone","https://www.sos.ms.gov/elections-voting"],["Sunflower","https://www.sos.ms.gov/elections-voting"],["Tallahatchie","https://www.sos.ms.gov/elections-voting"],["Tate","https://www.sos.ms.gov/elections-voting"],["Tippah","https://www.sos.ms.gov/elections-voting"],["Tishomingo","https://www.sos.ms.gov/elections-voting"],["Tunica","https://www.sos.ms.gov/elections-voting"],["Union","https://www.sos.ms.gov/elections-voting"],["Walthall","https://www.sos.ms.gov/elections-voting"],["Warren","https://www.sos.ms.gov/elections-voting"],["Washington","https://www.sos.ms.gov/elections-voting"],["Wayne","https://www.sos.ms.gov/elections-voting"],["Webster","https://www.sos.ms.gov/elections-voting"],["Wilkinson","https://www.sos.ms.gov/elections-voting"],["Winston","https://www.sos.ms.gov/elections-voting"],["Yalobusha","https://www.sos.ms.gov/elections-voting"],["Yazoo","https://www.sos.ms.gov/elections-voting"]],
gap:"Only federal races in 2026. Local/county/school board on different cycle.",
schoolNote:"Data held by county election commissions (when on-cycle).",
notes:"Primary held March 10. Only federal races in 2026."},
{s:"MO",n:"Missouri",filing:"Mar 31, 2026",primary:"Aug 4, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:false,govOpen:false,govNote:"",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Legislature"],
centralDb:{name:"MO Secretary of State",url:"https://www.sos.mo.gov/elections/s_default/candidate_info",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Clerk / Election Board",localCount:115,localDir:"https://www.sos.mo.gov/elections/sos_default/county_clerk_info",
counties:[["Adair","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Andrew","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Atchison","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Audrain","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Barry","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Barton","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Bates","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Benton","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Bollinger","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Boone","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Buchanan","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Butler","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Caldwell","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Callaway","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Camden","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Cape Girardeau","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Carroll","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Carter","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Cass","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Cedar","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Chariton","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Christian","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Clark","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Clay","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Clinton","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Cole","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Cooper","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Crawford","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Dade","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Dallas","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Daviess","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["DeKalb","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Dent","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Douglas","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Dunklin","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Franklin","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Gasconade","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Gentry","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Greene","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Grundy","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Harrison","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Henry","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Hickory","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Holt","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Howard","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Howell","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Iron","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Jackson","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Jasper","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Jefferson","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Johnson","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Knox","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Laclede","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Lafayette","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Lawrence","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Lewis","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Lincoln","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Linn","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Livingston","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Macon","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Madison","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Maries","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Marion","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["McDonald","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Mercer","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Miller","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Mississippi","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Moniteau","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Monroe","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Montgomery","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Morgan","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["New Madrid","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Newton","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Nodaway","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Oregon","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Osage","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Ozark","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Pemiscot","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Perry","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Pettis","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Phelps","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Pike","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Platte","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Polk","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Pulaski","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Putnam","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Ralls","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Randolph","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Ray","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Reynolds","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Ripley","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Saline","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Schuyler","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Scotland","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Scott","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Shannon","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Shelby","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["St. Charles","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["St. Clair","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["St. Francois","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["St. Louis City","https://www.stlouis-mo.gov/government/departments/board-election-commissioners/"],["St. Louis County","https://stlouiscountyelections.com/"],["Ste. Genevieve","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Stoddard","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Stone","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Sullivan","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Taney","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Texas","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Vernon","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Warren","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Washington","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Wayne","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Webster","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Worth","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"],["Wright","https://www.sos.mo.gov/elections/sos_default/county_clerk_info"]],
gap:"County, municipal, school board at county level. School boards often April elections.",
schoolNote:"Data held by school district secretaries or county clerks. April elections.",
notes:"No Governor or Senate race. House and legislature only."},
{s:"MT",n:"Montana",filing:"Mar 4, 2026",primary:"Jun 2, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:false,govOpen:false,govNote:"",senate:true,senOpen:true,senNote:"Open — Daines became RNC Chair",house:true,
otherOffices:["Legislature"],
centralDb:{name:"MT Secretary of State",url:"https://sosmt.gov/elections/filing/",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Election Administrator",localCount:56,localDir:"https://sosmt.gov/elections/resources/",
counties:[["Beaverhead","https://sosmt.gov/elections/resources/"],["Big Horn","https://sosmt.gov/elections/resources/"],["Blaine","https://sosmt.gov/elections/resources/"],["Broadwater","https://sosmt.gov/elections/resources/"],["Carbon","https://sosmt.gov/elections/resources/"],["Carter","https://sosmt.gov/elections/resources/"],["Cascade","https://sosmt.gov/elections/resources/"],["Chouteau","https://sosmt.gov/elections/resources/"],["Custer","https://sosmt.gov/elections/resources/"],["Daniels","https://sosmt.gov/elections/resources/"],["Dawson","https://sosmt.gov/elections/resources/"],["Deer Lodge","https://sosmt.gov/elections/resources/"],["Fallon","https://sosmt.gov/elections/resources/"],["Fergus","https://sosmt.gov/elections/resources/"],["Flathead","https://sosmt.gov/elections/resources/"],["Gallatin","https://sosmt.gov/elections/resources/"],["Garfield","https://sosmt.gov/elections/resources/"],["Glacier","https://sosmt.gov/elections/resources/"],["Golden Valley","https://sosmt.gov/elections/resources/"],["Granite","https://sosmt.gov/elections/resources/"],["Hill","https://sosmt.gov/elections/resources/"],["Jefferson","https://sosmt.gov/elections/resources/"],["Judith Basin","https://sosmt.gov/elections/resources/"],["Lake","https://sosmt.gov/elections/resources/"],["Lewis and Clark","https://sosmt.gov/elections/resources/"],["Liberty","https://sosmt.gov/elections/resources/"],["Lincoln","https://sosmt.gov/elections/resources/"],["Madison","https://sosmt.gov/elections/resources/"],["McCone","https://sosmt.gov/elections/resources/"],["Meagher","https://sosmt.gov/elections/resources/"],["Mineral","https://sosmt.gov/elections/resources/"],["Missoula","https://sosmt.gov/elections/resources/"],["Musselshell","https://sosmt.gov/elections/resources/"],["Park","https://sosmt.gov/elections/resources/"],["Petroleum","https://sosmt.gov/elections/resources/"],["Phillips","https://sosmt.gov/elections/resources/"],["Pondera","https://sosmt.gov/elections/resources/"],["Powder River","https://sosmt.gov/elections/resources/"],["Powell","https://sosmt.gov/elections/resources/"],["Prairie","https://sosmt.gov/elections/resources/"],["Ravalli","https://sosmt.gov/elections/resources/"],["Richland","https://sosmt.gov/elections/resources/"],["Roosevelt","https://sosmt.gov/elections/resources/"],["Rosebud","https://sosmt.gov/elections/resources/"],["Sanders","https://sosmt.gov/elections/resources/"],["Sheridan","https://sosmt.gov/elections/resources/"],["Silver Bow","https://sosmt.gov/elections/resources/"],["Stillwater","https://sosmt.gov/elections/resources/"],["Sweet Grass","https://sosmt.gov/elections/resources/"],["Teton","https://sosmt.gov/elections/resources/"],["Toole","https://sosmt.gov/elections/resources/"],["Treasure","https://sosmt.gov/elections/resources/"],["Valley","https://sosmt.gov/elections/resources/"],["Wheatland","https://sosmt.gov/elections/resources/"],["Wibaux","https://sosmt.gov/elections/resources/"],["Yellowstone","https://sosmt.gov/elections/resources/"]],
gap:"County, municipal, school board at county level.",
schoolNote:"Data held by county election administrators. May elections.",
notes:"Daines left for RNC. Open Senate seat in increasingly red state."},
{s:"NE",n:"Nebraska",filing:"Mar 2, 2026",primary:"May 12, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:false,govNote:"Pillen (R)",senate:true,senOpen:false,senNote:"Ricketts (R)",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Auditor","Unicameral Legislature"],
centralDb:{name:"NE SoS Candidate List",url:"https://sos.nebraska.gov/elections",covers:"Federal, statewide, legislature",format:"Searchable list",bulk:false},
localType:"County Clerk / Election Commissioner",localCount:93,localDir:"https://sos.nebraska.gov/elections/election-officials-contact-information",
counties:[["Adams","https://sos.nebraska.gov/elections"],["Antelope","https://sos.nebraska.gov/elections"],["Arthur","https://sos.nebraska.gov/elections"],["Banner","https://sos.nebraska.gov/elections"],["Blaine","https://sos.nebraska.gov/elections"],["Boone","https://sos.nebraska.gov/elections"],["Box Butte","https://sos.nebraska.gov/elections"],["Boyd","https://sos.nebraska.gov/elections"],["Brown","https://sos.nebraska.gov/elections"],["Buffalo","https://sos.nebraska.gov/elections"],["Burt","https://sos.nebraska.gov/elections"],["Butler","https://sos.nebraska.gov/elections"],["Cass","https://sos.nebraska.gov/elections"],["Cedar","https://sos.nebraska.gov/elections"],["Chase","https://sos.nebraska.gov/elections"],["Cherry","https://sos.nebraska.gov/elections"],["Cheyenne","https://sos.nebraska.gov/elections"],["Clay","https://sos.nebraska.gov/elections"],["Colfax","https://sos.nebraska.gov/elections"],["Cuming","https://sos.nebraska.gov/elections"],["Custer","https://sos.nebraska.gov/elections"],["Dakota","https://sos.nebraska.gov/elections"],["Dawes","https://sos.nebraska.gov/elections"],["Dawson","https://sos.nebraska.gov/elections"],["Deuel","https://sos.nebraska.gov/elections"],["Dixon","https://sos.nebraska.gov/elections"],["Dodge","https://sos.nebraska.gov/elections"],["Douglas","https://sos.nebraska.gov/elections"],["Dundy","https://sos.nebraska.gov/elections"],["Fillmore","https://sos.nebraska.gov/elections"],["Franklin","https://sos.nebraska.gov/elections"],["Frontier","https://sos.nebraska.gov/elections"],["Furnas","https://sos.nebraska.gov/elections"],["Gage","https://sos.nebraska.gov/elections"],["Garden","https://sos.nebraska.gov/elections"],["Garfield","https://sos.nebraska.gov/elections"],["Gosper","https://sos.nebraska.gov/elections"],["Grant","https://sos.nebraska.gov/elections"],["Greeley","https://sos.nebraska.gov/elections"],["Hall","https://sos.nebraska.gov/elections"],["Hamilton","https://sos.nebraska.gov/elections"],["Harlan","https://sos.nebraska.gov/elections"],["Hayes","https://sos.nebraska.gov/elections"],["Hitchcock","https://sos.nebraska.gov/elections"],["Holt","https://sos.nebraska.gov/elections"],["Hooker","https://sos.nebraska.gov/elections"],["Howard","https://sos.nebraska.gov/elections"],["Jefferson","https://sos.nebraska.gov/elections"],["Johnson","https://sos.nebraska.gov/elections"],["Kearney","https://sos.nebraska.gov/elections"],["Keith","https://sos.nebraska.gov/elections"],["Keya Paha","https://sos.nebraska.gov/elections"],["Kimball","https://sos.nebraska.gov/elections"],["Knox","https://sos.nebraska.gov/elections"],["Lancaster","https://sos.nebraska.gov/elections"],["Lincoln","https://sos.nebraska.gov/elections"],["Logan","https://sos.nebraska.gov/elections"],["Loup","https://sos.nebraska.gov/elections"],["Madison","https://sos.nebraska.gov/elections"],["McPherson","https://sos.nebraska.gov/elections"],["Merrick","https://sos.nebraska.gov/elections"],["Morrill","https://sos.nebraska.gov/elections"],["Nance","https://sos.nebraska.gov/elections"],["Nemaha","https://sos.nebraska.gov/elections"],["Nuckolls","https://sos.nebraska.gov/elections"],["Otoe","https://sos.nebraska.gov/elections"],["Pawnee","https://sos.nebraska.gov/elections"],["Perkins","https://sos.nebraska.gov/elections"],["Phelps","https://sos.nebraska.gov/elections"],["Pierce","https://sos.nebraska.gov/elections"],["Platte","https://sos.nebraska.gov/elections"],["Polk","https://sos.nebraska.gov/elections"],["Red Willow","https://sos.nebraska.gov/elections"],["Richardson","https://sos.nebraska.gov/elections"],["Rock","https://sos.nebraska.gov/elections"],["Saline","https://sos.nebraska.gov/elections"],["Sarpy","https://sos.nebraska.gov/elections"],["Saunders","https://sos.nebraska.gov/elections"],["Scotts Bluff","https://sos.nebraska.gov/elections"],["Seward","https://sos.nebraska.gov/elections"],["Sheridan","https://sos.nebraska.gov/elections"],["Sherman","https://sos.nebraska.gov/elections"],["Sioux","https://sos.nebraska.gov/elections"],["Stanton","https://sos.nebraska.gov/elections"],["Thayer","https://sos.nebraska.gov/elections"],["Thomas","https://sos.nebraska.gov/elections"],["Thurston","https://sos.nebraska.gov/elections"],["Valley","https://sos.nebraska.gov/elections"],["Washington","https://sos.nebraska.gov/elections"],["Wayne","https://sos.nebraska.gov/elections"],["Webster","https://sos.nebraska.gov/elections"],["Wheeler","https://sos.nebraska.gov/elections"],["York","https://sos.nebraska.gov/elections"]],
gap:"County, municipal, school board at county clerk level.",
schoolNote:"Data held by county clerks.",
notes:"Only state with unicameral nonpartisan legislature."},
{s:"NV",n:"Nevada",filing:"Mar 13, 2026",primary:"Jun 9, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:false,govNote:"Lombardo (R)",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Controller","Legislature"],
centralDb:{name:"NV Secretary of State",url:"https://www.nvsos.gov/sos/elections/election-information/candidates",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Clerk / Registrar",localCount:17,localDir:"https://www.nvsos.gov/sos/elections/voters/county-clerk-contact-information",
counties:[["Carson City","https://www.carson.org/government/departments-a-f/clerk-recorder/elections"],["Churchill","https://www.churchillcountynv.gov/191/Elections"],["Clark","https://www.clarkcountynv.gov/government/departments/elections/"],["Douglas","https://www.douglascountynv.gov/government/elected_officials/clerk_treasurer/elections_and_voting"],["Elko","https://www.elkocountynv.net/departments/clerk/elections.php"],["Esmeralda","https://www.nvsos.gov/sos/elections/voters/county-clerk-contact-information"],["Eureka","https://www.nvsos.gov/sos/elections/voters/county-clerk-contact-information"],["Humboldt","https://www.hcnv.us/178/Elections"],["Lander","https://www.nvsos.gov/sos/elections/voters/county-clerk-contact-information"],["Lincoln","https://www.nvsos.gov/sos/elections/voters/county-clerk-contact-information"],["Lyon","https://www.lyon-county.org/165/Elections"],["Mineral","https://www.nvsos.gov/sos/elections/voters/county-clerk-contact-information"],["Nye","https://www.nyecounty.net/179/Elections"],["Pershing","https://www.nvsos.gov/sos/elections/voters/county-clerk-contact-information"],["Storey","https://www.nvsos.gov/sos/elections/voters/county-clerk-contact-information"],["Washoe","https://www.washoecounty.gov/voters/"],["White Pine","https://www.nvsos.gov/sos/elections/voters/county-clerk-contact-information"]],
gap:"Only 17 counties — very manageable. Clark + Washoe = 90%+ of population.",
schoolNote:"Data held by county clerks.",
notes:"No Senate race. Lombardo (R) running for re-election in swing state. Only 17 counties."},
{s:"NH",n:"New Hampshire",filing:"Jun 12, 2026",primary:"Sep 8, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:false,govNote:"2-yr term",senate:true,senOpen:true,senNote:"Open — Shaheen retiring",house:true,
otherOffices:["Full Legislature (400-member House)"],
centralDb:{name:"NH Secretary of State",url:"https://sos.nh.gov/elections/candidates/",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"City/Town Clerk",localCount:234,localDir:"https://sos.nh.gov/elections/elections/election-officials/",
counties:[["Belknap","https://sos.nh.gov/elections/elections/election-officials/"],["Carroll","https://sos.nh.gov/elections/elections/election-officials/"],["Cheshire","https://sos.nh.gov/elections/elections/election-officials/"],["Coos","https://sos.nh.gov/elections/elections/election-officials/"],["Grafton","https://sos.nh.gov/elections/elections/election-officials/"],["Hillsborough","https://sos.nh.gov/elections/elections/election-officials/"],["Merrimack","https://sos.nh.gov/elections/elections/election-officials/"],["Rockingham","https://sos.nh.gov/elections/elections/election-officials/"],["Strafford","https://sos.nh.gov/elections/elections/election-officials/"],["Sullivan","https://sos.nh.gov/elections/elections/election-officials/"]],
gap:"All local/school board at town level. 234 municipalities.",
schoolNote:"Often elected at annual school district meetings. Data held by school district clerks.",
notes:"Governor every 2 years. Shaheen retiring. Largest state legislature (400 members)."},
{s:"NJ",n:"New Jersey",filing:"Mar 23, 2026",primary:"Jun 2, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:false,govOpen:false,govNote:"",senate:true,senOpen:false,senNote:"Booker (D)",house:true,
otherOffices:["No legislative races (off-cycle)"],
centralDb:{name:"NJ Division of Elections",url:"https://nj.gov/state/elections/candidate-information.shtml",covers:"Federal, statewide",format:"Online listings",bulk:false},
localType:"County Clerk",localCount:21,localDir:"https://www.nj.gov/state/elections/vote-county-election-officials.shtml",
counties:[["Atlantic","https://www.atlantic-county.org/elections/"],["Bergen","https://www.co.bergen.nj.us/county-clerk/elections"],["Burlington","https://www.co.burlington.nj.us/120/Elections"],["Camden","https://www.camdencounty.com/service/elections/"],["Cape May","https://www.capemaycountynj.gov/239/Board-of-Elections"],["Cumberland","https://www.cumberlandcountynj.gov/content/184/225/default.aspx"],["Essex","https://www.essexclerk.com/elections/"],["Gloucester","https://www.gloucestercountynj.gov/depts/c/countyclerk/elections/"],["Hudson","https://www.hudsoncountynj.org/county-clerk-elections"],["Hunterdon","https://www.co.hunterdon.nj.us/clerk/elections.html"],["Mercer","https://www.mercercounty.org/government/county-clerk/elections-information"],["Middlesex","https://www.middlesexcountynj.gov/government/departments/department-of-community-services/office-of-elections"],["Monmouth","https://www.monmouthcountyclerk.com/elections/"],["Morris","https://morriscountyclerk.org/elections/"],["Ocean","https://www.co.ocean.nj.us/cl/ElectionsInformation"],["Passaic","https://www.passaiccountynj.org/government/constitutional_officers/county_clerk/elections/"],["Salem","https://www.salemcountynj.gov/clerk/elections"],["Somerset","https://www.co.somerset.nj.us/government/elected-officials/county-clerk/elections"],["Sussex","https://www.sussex.nj.us/cn/webpage/5637/elections"],["Union","https://ucnj.org/county-clerk/elections/"],["Warren","https://www.co.warren.nj.us/Clerksite/CountyClerkElection.html"]],
gap:"Municipal and school board at county clerk level. Only 21 counties.",
schoolNote:"Data held by county clerks. November elections since 2012.",
notes:"No Governor or legislature. Senate and House only. Only 21 counties."},
{s:"NM",n:"New Mexico",filing:"Feb 3 / Mar 10, 2026",primary:"Jun 2, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:true,govNote:"Grisham term-limited",senate:true,senOpen:false,senNote:"Luján (D)",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Auditor","Comm. of Public Lands","House-only Legislature"],
centralDb:{name:"NM SoS Candidate Portal",url:"https://candidateportal.servis.sos.state.nm.us/CandidateList.aspx?eid=2911&cty=99",covers:"Federal, statewide, legislature",format:"Searchable portal",bulk:false},
localType:"County Clerk",localCount:33,localDir:"https://sos.state.nm.us/voting-and-elections/county-clerk-information/",
counties:[["Bernalillo","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Catron","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Chaves","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Cibola","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Colfax","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Curry","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["De Baca","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Dona Ana","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Eddy","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Grant","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Guadalupe","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Harding","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Hidalgo","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Lea","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Lincoln","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Los Alamos","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Luna","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["McKinley","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Mora","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Otero","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Quay","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Rio Arriba","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Roosevelt","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["San Juan","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["San Miguel","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Sandoval","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Santa Fe","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Sierra","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Socorro","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Taos","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Torrance","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Union","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"],["Valencia","https://sos.state.nm.us/voting-and-elections/county-clerk-information/"]],
gap:"County, municipal, school board at county clerk level. Only 33 counties.",
schoolNote:"Data held by county clerks.",
notes:"Grisham term-limited. 20% convention threshold. Only 33 counties."},
{s:"NY",n:"New York",filing:"Apr 6, 2026",primary:"Jun 23, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:false,govNote:"Hochul (D)",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Attorney General","Comptroller","Full Legislature"],
centralDb:{name:"NY State Board of Elections",url:"https://elections.ny.gov/filing-calendars",covers:"Federal, statewide, legislature",format:"Online listings",bulk:true},
localType:"County Board of Elections",localCount:62,localDir:"https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster",
counties:[["Albany","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Allegany","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Bronx","https://vote.nyc/"],["Broome","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Cattaraugus","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Cayuga","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Chautauqua","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Chemung","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Chenango","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Clinton","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Columbia","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Cortland","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Delaware","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Dutchess","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Erie","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Essex","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Franklin","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Fulton","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Genesee","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Greene","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Hamilton","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Herkimer","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Jefferson","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Kings","https://vote.nyc/"],["Lewis","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Livingston","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Madison","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Monroe","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Montgomery","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Nassau","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["New York","https://vote.nyc/"],["Niagara","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Oneida","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Onondaga","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Ontario","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Orange","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Orleans","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Oswego","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Otsego","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Putnam","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Queens","https://vote.nyc/"],["Rensselaer","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Richmond","https://vote.nyc/"],["Rockland","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Saratoga","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Schenectady","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Schoharie","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Schuyler","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Seneca","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["St. Lawrence","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Steuben","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Suffolk","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Sullivan","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Tioga","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Tompkins","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Ulster","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Warren","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Washington","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Wayne","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Westchester","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Wyoming","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"],["Yates","https://publicreporting.elections.ny.gov/CountyBoardRoster/CountyBoardRoster"]],
gap:"County, municipal, school board split across 62 BOEs + NYC BOE + school districts.",
schoolNote:"File with school district clerks. Many May elections (separate). NOT in any central system.",
notes:"No Senate. Hochul (D) re-election. 62 county BOEs + NYC BOE separate system (vote.nyc)."},
{s:"NC",n:"North Carolina",filing:"Dec 19, 2025",primary:"Mar 3, 2026",runoff:"May 12 (if requested)",general:"Nov 3, 2026",status:"voted",
gov:false,govOpen:false,govNote:"",senate:true,senOpen:true,senNote:"Open — Tillis not running; Cooper (D) vs Whatley (R)",house:true,
otherOffices:["Legislature"],
centralDb:{name:"NCSBE Candidate Lists",url:"https://www.ncsbe.gov/results-data/candidate-lists",covers:"★ ALL levels: federal, state, county, municipal, school board, judicial, special district",format:"★ Bulk CSV downloads on AWS S3",bulk:true},
localType:"N/A — Fully centralized",localCount:0,localDir:"https://www.ncsbe.gov/results-data/candidate-lists",
gap:"NONE — NC has the most data-friendly system in the US.",
schoolNote:"★ School board candidates ARE in the NCSBE bulk CSV downloads.",
notes:"★ GOLD STANDARD for data. Primary held Mar 3. Bulk CSV/S3. Cooper won Dem; Whatley won GOP."},
{s:"ND",n:"North Dakota",filing:"Apr 6, 2026",primary:"Jun 9, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:false,govOpen:false,govNote:"",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Attorney General","Secretary of State","Auditor","Ag Commissioner","Tax Commissioner","Insurance Commissioner","Legislature"],
centralDb:{name:"ND SoS Candidate Info",url:"https://vip.sos.nd.gov/candidateinfo.aspx",covers:"Federal, statewide, legislature",format:"Online search",bulk:false},
localType:"County Auditor",localCount:53,localDir:"https://vip.sos.nd.gov/CountyAuditors.aspx",
counties:[["Adams","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Barnes","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Benson","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Billings","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Bottineau","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Bowman","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Burke","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Burleigh","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Cass","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Cavalier","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Dickey","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Divide","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Dunn","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Eddy","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Emmons","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Foster","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Golden Valley","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Grand Forks","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Grant","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Griggs","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Hettinger","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Kidder","https://vip.sos.nd.gov/CountyAuditors.aspx"],["LaMoure","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Logan","https://vip.sos.nd.gov/CountyAuditors.aspx"],["McHenry","https://vip.sos.nd.gov/CountyAuditors.aspx"],["McIntosh","https://vip.sos.nd.gov/CountyAuditors.aspx"],["McKenzie","https://vip.sos.nd.gov/CountyAuditors.aspx"],["McLean","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Mercer","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Morton","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Mountrail","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Nelson","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Oliver","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Pembina","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Pierce","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Ramsey","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Ransom","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Renville","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Richland","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Rolette","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Sargent","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Sheridan","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Sioux","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Slope","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Stark","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Steele","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Stutsman","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Towner","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Traill","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Walsh","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Ward","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Wells","https://vip.sos.nd.gov/CountyAuditors.aspx"],["Williams","https://vip.sos.nd.gov/CountyAuditors.aspx"]],
gap:"County, municipal, school board at county auditor level.",
schoolNote:"Data held by county auditors or school business managers.",
notes:"No Governor or Senate. No voter registration (only state without it)."},
{s:"OH",n:"Ohio",filing:"Feb 4, 2026",primary:"May 5, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:true,govNote:"DeWine term-limited",senate:true,senOpen:false,senNote:"Special — Vance resigned (VP); Husted appointed",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Auditor","Legislature"],
centralDb:{name:"OH Secretary of State",url:"https://www.ohiosos.gov/elections/elections-officials/2026-elections/",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Board of Elections",localCount:88,localDir:"https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/",
counties:[["Adams","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Allen","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Ashland","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Ashtabula","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Athens","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Auglaize","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Belmont","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Brown","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Butler","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Carroll","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Champaign","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Clark","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Clermont","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Clinton","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Columbiana","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Coshocton","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Crawford","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Cuyahoga","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Darke","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Defiance","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Delaware","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Erie","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Fairfield","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Fayette","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Franklin","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Fulton","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Gallia","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Geauga","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Greene","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Guernsey","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Hamilton","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Hancock","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Hardin","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Harrison","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Henry","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Highland","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Hocking","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Holmes","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Huron","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Jackson","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Jefferson","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Knox","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Lake","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Lawrence","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Licking","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Logan","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Lorain","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Lucas","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Madison","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Mahoning","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Marion","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Medina","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Meigs","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Mercer","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Miami","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Monroe","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Montgomery","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Morgan","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Morrow","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Muskingum","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Noble","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Ottawa","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Paulding","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Perry","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Pickaway","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Pike","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Portage","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Preble","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Putnam","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Richland","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Ross","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Sandusky","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Scioto","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Seneca","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Shelby","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Stark","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Summit","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Trumbull","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Tuscarawas","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Union","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Van Wert","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Vinton","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Warren","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Washington","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Wayne","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Williams","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Wood","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"],["Wyandot","https://www.ohiosos.gov/elections/elections-officials/county-boards-of-elections-directory/"]],
gap:"County, municipal, school board, special district at county BOE level. 88 BOEs.",
schoolNote:"Data held by county boards of elections.",
notes:"DeWine term-limited. Special Senate for VP Vance's seat. DATA Act at data.ohiosos.gov."},
{s:"OK",n:"Oklahoma",filing:"Apr 3, 2026",primary:"Jun 16, 2026",runoff:"Aug 25, 2026",general:"Nov 3, 2026",status:"open",
gov:true,govOpen:true,govNote:"Stitt term-limited",senate:true,senOpen:false,senNote:"Mullin (R)",house:true,
otherOffices:["Attorney General","Auditor","Corporation Commissioner","Insurance Commissioner","Legislature"],
centralDb:{name:"OK State Election Board",url:"https://oklahoma.gov/elections/candidates.html",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Election Board",localCount:77,localDir:"https://oklahoma.gov/elections/about-us/county-election-boards.html",
counties:[["Adair","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Alfalfa","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Atoka","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Beaver","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Beckham","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Blaine","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Bryan","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Caddo","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Canadian","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Carter","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Cherokee","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Choctaw","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Cimarron","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Cleveland","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Coal","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Comanche","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Cotton","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Craig","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Creek","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Custer","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Delaware","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Dewey","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Ellis","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Garfield","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Garvin","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Grady","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Grant","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Greer","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Harmon","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Harper","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Haskell","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Hughes","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Jackson","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Jefferson","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Johnston","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Kay","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Kingfisher","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Kiowa","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Latimer","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Le Flore","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Lincoln","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Logan","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Love","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Major","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Marshall","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Mayes","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["McClain","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["McCurtain","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["McIntosh","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Murray","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Muskogee","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Noble","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Nowata","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Okfuskee","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Oklahoma","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Okmulgee","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Osage","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Ottawa","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Pawnee","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Payne","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Pittsburg","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Pontotoc","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Pottawatomie","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Pushmataha","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Roger Mills","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Rogers","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Seminole","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Sequoyah","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Stephens","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Texas","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Tillman","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Tulsa","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Wagoner","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Washington","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Washita","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Woods","https://oklahoma.gov/elections/about-us/county-election-boards.html"],["Woodward","https://oklahoma.gov/elections/about-us/county-election-boards.html"]],
gap:"County, municipal, school board at county election board level.",
schoolNote:"Data held by county election boards. February elections (separate cycle).",
notes:"Stitt term-limited. Latest runoff in country (Aug 25)."},
{s:"OR",n:"Oregon",filing:"Mar 10, 2026",primary:"May 19, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:false,govNote:"Kotek (D)",senate:true,senOpen:false,senNote:"Merkley (D)",house:true,
otherOffices:["Treasurer","Legislature"],
centralDb:{name:"OR SoS + ORESTAR",url:"https://sos.oregon.gov/elections/Pages/Candidate-Filings-Local-Measures.aspx",covers:"Federal, statewide, legislature",format:"Online listings + ORESTAR API",bulk:true},
localType:"County Clerk",localCount:36,localDir:"https://sos.oregon.gov/elections/Pages/countyofficials.aspx",
counties:[["Baker","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Benton","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Clackamas","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Clatsop","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Columbia","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Coos","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Crook","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Curry","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Deschutes","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Douglas","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Gilliam","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Grant","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Harney","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Hood River","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Jackson","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Jefferson","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Josephine","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Klamath","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Lake","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Lane","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Lincoln","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Linn","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Malheur","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Marion","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Morrow","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Multnomah","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Polk","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Sherman","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Tillamook","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Umatilla","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Union","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Wallowa","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Wasco","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Washington","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Wheeler","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"],["Yamhill","https://sos.oregon.gov/elections/Pages/countyofficials.aspx"]],
gap:"County, municipal, school board at county clerk level. Only 36 counties.",
schoolNote:"Data held by county clerks.",
notes:"All-mail voting state. ORESTAR has API for campaign finance. Only 36 counties."},
{s:"PA",n:"Pennsylvania",filing:"Mar 10, 2026",primary:"May 19, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:false,govNote:"Shapiro (D)",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Treasurer","Auditor General","Legislature"],
centralDb:{name:"PA Department of State",url:"https://www.dos.pa.gov/VotingElections/CandidatesCommittees/Pages/default.aspx",covers:"Federal, statewide, legislature (LIMITED)",format:"Weak — primarily election results, not filing search",bulk:false},
localType:"County Board of Elections",localCount:67,localDir:"https://www.dos.pa.gov/VotingElections/OtherServicesEvents/Pages/County-Contact-Information.aspx",
// PA - Pennsylvania (67 counties)
counties:[
["Adams","https://www.adamscountypa.gov/elections"],
["Allegheny","https://www.alleghenycounty.us/elections"],
["Armstrong","https://www.co.armstrong.pa.us/departments/elections-voter-registration"],
["Beaver","https://www.beavercountypa.gov/depts/rov/pages/default.aspx"],
["Bedford","https://bedfordcountypa.org/elections/"],
["Berks","https://www.berkspa.gov/dept/elections"],
["Blair","https://www.blaircountypa.gov/281/Elections-Voter-Registration"],
["Bradford","https://www.bradfordcountypa.org/Elections-Voter-Registration"],
["Bucks","https://www.buckscounty.gov/1252/Board-of-Elections"],
["Butler","https://www.butlercountypa.gov/161/Bureau-of-Elections"],
["Cambria","https://www.cambriacountypa.gov/election-office/"],
["Cameron","https://www.cameroncountypa.com/elections"],
["Carbon","https://www.carboncounty.com/index.php/elections"],
["Centre","https://centrecountypa.gov/424/Elections-Voter-Registration"],
["Chester","https://www.chesco.org/154/Voter-Services"],
["Clarion","https://www.co.clarion.pa.us/departments/elections/index.php"],
["Clearfield","https://www.clearfieldco.org/index.php/elections"],
["Clinton","https://www.clintoncountypa.com/county-government/elections"],
["Columbia","https://www.columbiapa.org/election-bureau"],
["Crawford","https://www.crawfordcountypa.net/elections"],
["Cumberland","https://www.cumberlandcountypa.gov/1902/Voter-Registration"],
["Dauphin","https://www.dauphincounty.gov/government/support-services/registration-elections"],
["Delaware","https://www.delcopa.gov/electionsbureau/"],
["Elk","https://www.co.elk.pa.us/departments/elections"],
["Erie","https://eriecountypa.gov/departments/elections-and-voter-registration/"],
["Fayette","https://www.fayettecountypa.org/441/Elections"],
["Forest","https://www.co.forest.pa.us/departments/elections/index.htm"],
["Franklin","https://franklincountypa.gov/index.php/elections"],
["Fulton","https://www.fultoncountypa.gov/county-offices/elections"],
["Greene","https://www.co.greene.pa.us/departments/elections"],
["Huntingdon","https://huntingdoncounty.net/elections/"],
["Indiana","https://www.indianacountypa.gov/departments/elections/"],
["Jefferson","https://www.jeffersoncountypa.com/departments/elections-voter-registration"],
["Juniata","https://www.juniataco.org/voter-registration-elections/"],
["Lackawanna","https://www.lackawannacounty.org/index.php/election-bureau"],
["Lancaster","https://www.lancasterpa.gov/elections/"],
["Lawrence","https://www.lawrencecountypa.gov/departments/voter-registration-elections"],
["Lebanon","https://www.lebcounty.org/depts/Elections/Pages/default.aspx"],
["Lehigh","https://www.lehighcounty.org/departments/voter-registration-elections"],
["Luzerne","https://www.luzernecounty.org/841/Bureau-of-Elections"],
["Lycoming","https://www.lyco.org/Departments/Elections-Voter-Registration"],
["McKean","https://www.mckeancountypa.org/departments/elections"],
["Mercer","https://www.mercercountypa.gov/depts/elections/index.htm"],
["Mifflin","https://www.co.mifflin.pa.us/departments/elections/index.php"],
["Monroe","https://www.monroecountypa.gov/departments/elections-voter-registration"],
["Montgomery","https://www.montcopa.org/753/Voter-Services"],
["Montour","https://montourco.org/departments/elections/"],
["Northampton","https://www.northamptoncounty.org/CTYADMN/ELECTNS/Pages/default.aspx"],
["Northumberland","https://www.norrycopa.net/departments/elections-voter-registration"],
["Perry","https://www.perrycountypa.org/government/elections-voter-registration"],
["Philadelphia","https://vote.phila.gov/"],
["Pike","https://www.pikepa.org/government/elections_office/index.php"],
["Potter","https://www.pottercountypa.net/county-offices/elections-voter-registration"],
["Schuylkill","https://www.co.schuylkill.pa.us/Offices/ElectionBureau/ElectionBureau.asp"],
["Snyder","https://www.snydercounty.org/depts/elections/"],
["Somerset","https://www.co.somerset.pa.us/departments/elections-voter-registration"],
["Sullivan","https://www.sullivancounty-pa.us/departments/elections-voter-registration"],
["Susquehanna","https://www.susqco.com/elections.html"],
["Tioga","https://www.tiogacountypa.us/departments/elections-voter-registration"],
["Union","https://www.unionco.org/departments/elections-voter-registration"],
["Venango","https://www.co.venango.pa.us/230/Elections-Voter-Registration"],
["Warren","https://www.warren-county.net/departments/elections-voter-registration"],
["Washington","https://www.washingtoncopa.gov/elections"],
["Wayne","https://www.waynecountypa.gov/356/Elections-Voter-Registration"],
["Westmoreland","https://www.co.westmoreland.pa.us/245/Elections"],
["Wyoming","https://www.wycopa.org/departments/elections-voter-registration"],
["York","https://yorkcountypa.gov/voting-elections/"]
],
gap:"Almost everything below state legislature requires county boards. WEAKEST state system.",
schoolNote:"Data held by county boards. Partisan school board elections (unusual).",
notes:"No Senate. Shapiro (D) re-election. Key swing state. WORST state-level data system."},
{s:"RI",n:"Rhode Island",filing:"Jun 24, 2026",primary:"Sep 8, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:false,govNote:"McKee (D)",senate:true,senOpen:false,senNote:"Reed (D)",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Full Legislature"],
centralDb:{name:"RI Board of Elections",url:"https://elections.ri.gov/elections/upcoming-elections",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"Local Board of Canvassers",localCount:39,localDir:"https://elections.ri.gov/about/local-boards-canvassers",
// RI - Rhode Island (39 municipalities)
counties:[
["Barrington","https://www.barrington.ri.gov/boardofcanvassers"],
["Bristol","https://www.bristolri.gov/board-of-canvassers"],
["Burrillville","https://www.burrillville.org/board-canvassers"],
["Central Falls","https://www.centralfallsri.gov/board-of-canvassers"],
["Charlestown","https://www.charlestownri.org/board-canvassers"],
["Coventry","https://coventryri.gov/board-canvassers"],
["Cranston","https://www.cranstonri.gov/boards-commissions/board-of-canvassers/"],
["Cumberland","https://www.cumberlandri.gov/166/Board-of-Canvassers"],
["East Greenwich","https://www.eastgreenwichri.com/482/Board-of-Canvassers"],
["East Providence","https://www.eastprovidenceri.gov/board-canvassers"],
["Exeter","https://www.exeterri.gov/board-canvassers"],
["Foster","https://www.fostergov.com/board-of-canvassers"],
["Glocester","https://www.glocesterri.org/board-canvassers"],
["Hopkinton","https://www.hopkintonri.gov/board-of-canvassers"],
["Jamestown","https://www.jamestownri.gov/board-of-canvassers"],
["Johnston","https://www.townofjohnstonri.com/board-of-canvassers"],
["Lincoln","https://www.lincolnri.gov/board-of-canvassers"],
["Little Compton","https://www.littlecomptonri.org/board-of-canvassers"],
["Middletown","https://www.middletownri.com/154/Board-of-Canvassers"],
["Narragansett","https://www.narragansettri.gov/board-canvassers"],
["New Shoreham","https://www.new-shoreham.com/board-of-canvassers"],
["Newport","https://www.cityofnewport.com/departments/board-of-canvassers"],
["North Kingstown","https://www.northkingstown.org/301/Board-of-Canvassers"],
["North Providence","https://northprovidenceri.gov/board-of-canvassers/"],
["North Smithfield","https://www.nsmithfieldri.org/board-canvassers"],
["Pawtucket","https://www.pawtucketri.com/board-canvassers"],
["Portsmouth","https://www.portsmouthri.gov/231/Board-of-Canvassers"],
["Providence","https://www.providenceri.gov/board-of-canvassers/"],
["Richmond","https://www.richmondri.com/board-of-canvassers"],
["Scituate","https://www.scituateri.org/board-of-canvassers"],
["Smithfield","https://www.smithfieldri.com/board-canvassers"],
["South Kingstown","https://www.southkingstownri.com/260/Board-of-Canvassers"],
["Tiverton","https://www.tiverton.ri.gov/board-of-canvassers"],
["Warren","https://www.townofwarren-ri.gov/town_government/boards_and_commissions/board_of_canvassers.php"],
["Warwick","https://www.warwickri.gov/board-canvassers"],
["West Greenwich","https://www.wgreenwich.ri.gov/board-of-canvassers"],
["West Warwick","https://www.westwarwickri.org/board-of-canvassers"],
["Westerly","https://westerlyri.gov/306/Canvassers-Board-of"],
["Woonsocket","https://www.woonsocketri.org/board-canvassers"]
],
gap:"Municipal and school committee at local board level. Only 39 — manageable.",
schoolNote:"Called 'school committees.' Data held by local boards of canvassers.",
notes:"No county government. Only 39 municipalities."},
{s:"SC",n:"South Carolina",filing:"Mar 30, 2026",primary:"Jun 9, 2026",runoff:"Jun 23, 2026",general:"Nov 3, 2026",status:"open",
gov:true,govOpen:true,govNote:"McMaster term-limited",senate:true,senOpen:false,senNote:"Graham (R)",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Comptroller","Supt. of Education","House-only Legislature"],
centralDb:{name:"SC Election Commission Candidate Search",url:"https://scvotes.gov/candidates/",covers:"Federal, statewide, legislature, some county",format:"Searchable database",bulk:false},
localType:"County Voter Registration Office",localCount:46,localDir:"https://scvotes.gov/contact/county-voter-registration-election-offices/",
// SC - South Carolina (46 counties)
counties:[
["Abbeville","https://www.abbevillecountysc.com/voter-registration-elections"],
["Aiken","https://www.aikencountysc.gov/departments/voter-registration-elections/"],
["Allendale","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Anderson","https://www.andersoncountysc.org/departments-a-z/vote/"],
["Bamberg","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Barnwell","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Beaufort","https://www.beaufortcountysc.gov/vote/index.html"],
["Berkeley","https://berkeleycountysc.gov/dept/elections/"],
["Calhoun","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Charleston","https://www.charlestoncounty.org/departments/bevr/"],
["Cherokee","https://www.cherokeecountysc.com/voter-registration-and-elections"],
["Chester","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Chesterfield","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Clarendon","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Colleton","https://www.colletoncounty.org/voter-registration-elections"],
["Darlington","https://www.darcosc.com/departments/voter-registration-elections/"],
["Dillon","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Dorchester","https://www.dorchestercountysc.gov/government/voter-registration-and-elections"],
["Edgefield","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Fairfield","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Florence","https://www.florenceco.org/departments/voter-registration-elections"],
["Georgetown","https://www.georgetowncountysc.org/voter-registration/"],
["Greenville","https://www.greenvillecounty.org/VoterRegistration/"],
["Greenwood","https://www.greenwoodcounty-sc.gov/voter-registration"],
["Hampton","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Horry","https://www.horrycounty.org/departments/voter-registration"],
["Jasper","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Kershaw","https://www.kershaw.sc.gov/voter-registration-elections"],
["Lancaster","https://www.mylancastersc.org/voter-registration-elections"],
["Laurens","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Lee","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Lexington","https://lex-co.sc.gov/departments/voter-registration-elections"],
["Marion","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Marlboro","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["McCormick","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Newberry","https://www.newberrycounty.net/voter-registration-elections"],
["Oconee","https://www.oconeesc.com/Departments-Services/Voter-Registration-Elections"],
["Orangeburg","https://www.orangeburgcounty.org/voter-registration-elections/"],
["Pickens","https://www.co.pickens.sc.us/voter-registration-elections"],
["Richland","https://www.richlandcountysc.gov/Government/Departments/Board-of-Voter-Registration-Elections"],
["Saluda","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Spartanburg","https://www.spartanburgcounty.org/281/Voter-Registration-Elections"],
["Sumter","https://www.sumtercountysc.org/voter-registration-elections"],
["Union","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["Williamsburg","https://scvotes.gov/contact/county-voter-registration-election-offices/"],
["York","https://www.yorkcountygov.com/155/Board-of-Voter-Registration-Elections"]
],
gap:"Municipal, school board, some county at county office level.",
schoolNote:"Data held by county voter registration offices.",
notes:"McMaster term-limited. Filing closed Mar 30. Senate not up."},
{s:"SD",n:"South Dakota",filing:"Mar 31, 2026",primary:"Jun 2, 2026",runoff:"Jul 28, 2026",general:"Nov 3, 2026",status:"open",
gov:true,govOpen:true,govNote:"Noem left for DHS",senate:true,senOpen:false,senNote:"Rounds (R)",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Auditor","Comm. of School/Public Lands","Full Legislature"],
centralDb:{name:"SD Secretary of State",url:"https://sdsos.gov/elections-voting/upcoming-elections/candidate-information.aspx",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Auditor",localCount:66,localDir:"https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx",
// SD - South Dakota (66 counties)
counties:[
["Aurora","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Beadle","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Bennett","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Bon Homme","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Brookings","https://www.brookingscountysd.gov/152/Auditor"],
["Brown","https://brown.sd.us/auditor"],
["Brule","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Buffalo","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Butte","https://www.buttecountysd.org/auditor"],
["Campbell","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Charles Mix","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Clark","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Clay","https://www.claycountysd.org/auditor"],
["Codington","https://www.codington.org/auditor/"],
["Corson","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Custer","https://www.custercountysd.com/auditor"],
["Davison","https://www.davisoncounty.org/auditor"],
["Day","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Deuel","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Dewey","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Douglas","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Edmunds","https://edmunds.sdcounties.org/auditor/"],
["Fall River","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Faulk","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Grant","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Gregory","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Haakon","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Hamlin","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Hand","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Hanson","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Harding","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Hughes","https://hughescounty.sd.gov/auditor/"],
["Hutchinson","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Hyde","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Jackson","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Jerauld","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Jones","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Kingsbury","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Lake","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Lawrence","https://www.lawrence.sd.us/auditor/"],
["Lincoln","https://www.lincolncountysd.gov/158/Auditor"],
["Lyman","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Marshall","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["McCook","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["McPherson","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Meade","https://www.meadecounty.org/auditor"],
["Mellette","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Miner","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Minnehaha","https://www.minnehahacounty.gov/dept/au/au.php"],
["Moody","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Oglala Lakota","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Pennington","https://pennco.org/auditor"],
["Perkins","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Potter","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Roberts","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Sanborn","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Spink","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Stanley","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Sully","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Todd","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Tripp","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Turner","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Union","https://unioncountysd.org/auditor/"],
["Walworth","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"],
["Yankton","https://www.co.yankton.sd.us/auditor/"],
["Ziebach","https://sdsos.gov/elections-voting/upcoming-elections/county-auditors.aspx"]
],
gap:"County, municipal, school board at county auditor level.",
schoolNote:"Data held by school business managers or county auditors. June elections.",
notes:"Noem left for DHS Secretary — open seat."},
{s:"TN",n:"Tennessee",filing:"Mar 10, 2026",primary:"Aug 6, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:true,govOpen:true,govNote:"Lee term-limited",senate:true,senOpen:false,senNote:"Hagerty (R)",house:true,
otherOffices:["Legislature"],
centralDb:{name:"TN SoS Candidate Search",url:"https://sos.tn.gov/elections/2026-candidate-lists",covers:"Federal, statewide, legislature",format:"Searchable database",bulk:false},
localType:"County Election Commission",localCount:95,localDir:"https://tnsos.org/elections/election_commissions.php",
// TN - Tennessee (95 counties)
counties:[
["Anderson","https://andersoncountytn.gov/election-commission/"],
["Bedford","https://www.bedfordcountytn.gov/election-commission"],
["Benton","https://www.bentoncotnelections.com/"],
["Bledsoe","https://tnsos.org/elections/election_commissions.php"],
["Blount","https://www.blounttn.gov/447/Election-Commission"],
["Bradley","https://bradleyelections.com/"],
["Campbell","https://tnsos.org/elections/election_commissions.php"],
["Cannon","https://tnsos.org/elections/election_commissions.php"],
["Carroll","https://carrollcountytnelections.com/"],
["Carter","https://tnsos.org/elections/election_commissions.php"],
["Cheatham","https://www.cheathamcountytn.gov/election_commission.php"],
["Chester","https://tnsos.org/elections/election_commissions.php"],
["Claiborne","https://tnsos.org/elections/election_commissions.php"],
["Clay","https://tnsos.org/elections/election_commissions.php"],
["Cocke","https://tnsos.org/elections/election_commissions.php"],
["Coffee","https://www.coffeecountytn.gov/election-commission"],
["Crockett","https://tnsos.org/elections/election_commissions.php"],
["Cumberland","https://www.cumberlandcountytn.gov/election-commission"],
["Davidson","https://www.nashville.gov/departments/elections"],
["Decatur","https://tnsos.org/elections/election_commissions.php"],
["DeKalb","https://tnsos.org/elections/election_commissions.php"],
["Dickson","https://www.dicksoncountytn.gov/election_commission.php"],
["Dyer","https://tnsos.org/elections/election_commissions.php"],
["Fayette","https://tnsos.org/elections/election_commissions.php"],
["Fentress","https://tnsos.org/elections/election_commissions.php"],
["Franklin","https://tnsos.org/elections/election_commissions.php"],
["Gibson","https://tnsos.org/elections/election_commissions.php"],
["Giles","https://tnsos.org/elections/election_commissions.php"],
["Grainger","https://tnsos.org/elections/election_commissions.php"],
["Greene","https://tnsos.org/elections/election_commissions.php"],
["Grundy","https://tnsos.org/elections/election_commissions.php"],
["Hamblen","https://tnsos.org/elections/election_commissions.php"],
["Hamilton","https://elect.hamiltontn.gov/"],
["Hancock","https://tnsos.org/elections/election_commissions.php"],
["Hardeman","https://tnsos.org/elections/election_commissions.php"],
["Hardin","https://tnsos.org/elections/election_commissions.php"],
["Hawkins","https://tnsos.org/elections/election_commissions.php"],
["Haywood","https://tnsos.org/elections/election_commissions.php"],
["Henderson","https://tnsos.org/elections/election_commissions.php"],
["Henry","https://tnsos.org/elections/election_commissions.php"],
["Hickman","https://tnsos.org/elections/election_commissions.php"],
["Houston","https://tnsos.org/elections/election_commissions.php"],
["Humphreys","https://tnsos.org/elections/election_commissions.php"],
["Jackson","https://tnsos.org/elections/election_commissions.php"],
["Jefferson","https://tnsos.org/elections/election_commissions.php"],
["Johnson","https://tnsos.org/elections/election_commissions.php"],
["Knox","https://www.knoxcounty.org/election/"],
["Lake","https://tnsos.org/elections/election_commissions.php"],
["Lauderdale","https://tnsos.org/elections/election_commissions.php"],
["Lawrence","https://tnsos.org/elections/election_commissions.php"],
["Lewis","https://tnsos.org/elections/election_commissions.php"],
["Lincoln","https://tnsos.org/elections/election_commissions.php"],
["Loudon","https://tnsos.org/elections/election_commissions.php"],
["Macon","https://tnsos.org/elections/election_commissions.php"],
["Madison","https://www.madisoncountytn.gov/election-commission"],
["Marion","https://tnsos.org/elections/election_commissions.php"],
["Marshall","https://tnsos.org/elections/election_commissions.php"],
["Maury","https://www.maurycounty-tn.gov/227/Election-Commission"],
["McMinn","https://tnsos.org/elections/election_commissions.php"],
["McNairy","https://tnsos.org/elections/election_commissions.php"],
["Meigs","https://tnsos.org/elections/election_commissions.php"],
["Monroe","https://tnsos.org/elections/election_commissions.php"],
["Montgomery","https://mcgtn.org/election"],
["Moore","https://tnsos.org/elections/election_commissions.php"],
["Morgan","https://tnsos.org/elections/election_commissions.php"],
["Obion","https://tnsos.org/elections/election_commissions.php"],
["Overton","https://tnsos.org/elections/election_commissions.php"],
["Perry","https://tnsos.org/elections/election_commissions.php"],
["Pickett","https://tnsos.org/elections/election_commissions.php"],
["Polk","https://tnsos.org/elections/election_commissions.php"],
["Putnam","https://www.putnamcountytn.gov/election-commission"],
["Rhea","https://tnsos.org/elections/election_commissions.php"],
["Roane","https://tnsos.org/elections/election_commissions.php"],
["Robertson","https://www.robertsoncountytn.gov/election-commission"],
["Rutherford","https://rutherfordcountytn.gov/election-commission/"],
["Scott","https://tnsos.org/elections/election_commissions.php"],
["Sequatchie","https://tnsos.org/elections/election_commissions.php"],
["Sevier","https://www.seviercountytn.gov/election-commission"],
["Shelby","https://www.electioncommission.shelbycountytn.gov/"],
["Smith","https://tnsos.org/elections/election_commissions.php"],
["Stewart","https://tnsos.org/elections/election_commissions.php"],
["Sullivan","https://www.sullivancountytn.gov/election-commission"],
["Sumner","https://www.sumnertnelections.com/"],
["Tipton","https://tnsos.org/elections/election_commissions.php"],
["Trousdale","https://tnsos.org/elections/election_commissions.php"],
["Unicoi","https://tnsos.org/elections/election_commissions.php"],
["Union","https://tnsos.org/elections/election_commissions.php"],
["Van Buren","https://tnsos.org/elections/election_commissions.php"],
["Warren","https://tnsos.org/elections/election_commissions.php"],
["Washington","https://washingtoncountytn.gov/election-commission/"],
["Wayne","https://tnsos.org/elections/election_commissions.php"],
["Weakley","https://tnsos.org/elections/election_commissions.php"],
["White","https://tnsos.org/elections/election_commissions.php"],
["Williamson","https://www.williamsoncounty-tn.gov/40/Election-Commission"],
["Wilson","https://www.wilsoncountytn.gov/election-commission"]
],
gap:"County, municipal, school board at county election commission level.",
schoolNote:"Data held by county election commissions.",
notes:"Lee term-limited. Longest filing-to-primary gap (149 days)."},
{s:"TX",n:"Texas",filing:"Dec 8, 2025",primary:"Mar 3, 2026",runoff:"May 26, 2026",general:"Nov 3, 2026",status:"voted",
gov:true,govOpen:false,govNote:"Abbott (R) renominated",senate:true,senOpen:false,senNote:"Cornyn vs Paxton runoff May 26",house:true,
otherOffices:["Attorney General","Comptroller","Ag Commissioner","Land Commissioner","Legislature"],
centralDb:{name:"TX Secretary of State",url:"https://www.sos.state.tx.us/elections/candidates/index.shtml",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Elections Administrator / County Clerk",localCount:254,localDir:"https://www.sos.texas.gov/elections/voter/county.shtml",
counties:[["Harris","https://www.harrisvotes.com/"],["Dallas","https://www.dallascountyvotes.org/"],["Tarrant","https://www.tarrantcountyelections.com/"],["Bexar","https://www.bexar.org/1568/Elections"],["Travis","https://www.traviscountyelections.org/"],["Collin","https://www.collincountytx.gov/elections"],["Hidalgo","https://www.hidalgocounty.us/171/Elections-Department"],["El Paso","https://www.epcountyvotes.com/"],["Denton","https://www.votedenton.gov/"],["Fort Bend","https://www.fortbendcountytx.gov/government/departments/elections"],["Williamson","https://www.wilco.org/elections"],["Montgomery","https://www.mctx.org/departments/elections/"],["Nueces","https://www.nuecesco.com/county-services/elections-administration"],["Lubbock","https://www.mylubbock.us/departmental-websites/departments/elections/elections-home"],["Cameron","https://www.cameroncountytx.gov/elections/"],["Webb","https://www.webbcountytx.gov/Elections/"],["Bell","https://www.bellcountytx.com/county_government/elections/"],["Galveston","https://www.galvestoncountytx.gov/our-government/county-judge/elections"],["Brazoria","https://www.brazoriacountytx.gov/departments/elections-administration"],["Smith","https://www.smith-county.com/government/departments/elections"],["McLennan","https://www.mclennan.vote/"],["Hays","https://hayscountytx.com/departments/elections-administrator/"],["Jefferson","https://www.co.jefferson.tx.us/elections/"],["Brazos","https://www.brazoscountyelections.com/"],["Johnson","https://www.johnsoncountytx.org/departments/elections"],["Guadalupe","https://www.guadalupecountytx.us/elections/"],["Midland","https://www.co.midland.tx.us/222/Elections"],["Ector","https://www.co.ector.tx.us/page/Elections.Department"],["Comal","https://www.comalcountytx.com/departments/elections/"],["Kaufman","https://www.kaufmancounty.net/elections"],["Wichita","https://www.co.wichita.tx.us/Elections/"],["Potter","https://www.randallcounty.com/elections/"],["Ellis","https://co.ellis.tx.us/348/Elections-Administrator"],["Tom Green","https://www.tomgreencountytx.gov/page/tgc.Elections"],["Parker","https://www.parkercountytx.com/1173/Elections"],["Bowie","https://www.txkusa.org/129/Elections"],["Taylor","https://www.taylorcountytexas.org/249/Elections-Department"],["Grayson","https://www.co.grayson.tx.us/page/grayson.Elections"],["Randall","https://www.randallcounty.com/elections/"],["Rockwall","https://www.rockwallcountytexas.com/345/Elections"],["Liberty","https://www.co.liberty.tx.us/page/liberty.Elections"],["Orange","https://www.co.orange.tx.us/page/OC.Elections"],["Hunt","https://www.huntcountytx.net/elections"],["Victoria","https://www.vctx.org/government/county_offices/elections/"],["Henderson","https://www.hendersoncountytx.us/elections/"],["Angelina","https://www.angelinacounty.net/elections"],["Gregg","https://www.co.gregg.tx.us/page/gregg.Elections"],["San Patricio","https://www.co.san-patricio.tx.us/page/sanpatricio.Elections"],["Bastrop","https://www.co.bastrop.tx.us/page/bastrop.Elections.Department"],["Wise","https://www.co.wise.tx.us/elections/"],["Cherokee","https://www.cherokeecountytx.us/page/cherokee.Elections"],["Starr","https://www.co.starr.tx.us/page/StarrCounty.Elections"]],
gap:"HARDEST STATE — 254 counties + ~1,200 school district secretaries + thousands of special districts. Local elections in May.",
schoolNote:"Data held by SCHOOL BOARD SECRETARIES (not county). ~1,200 school districts. May elections. No central database.",
notes:"Primary held Mar 3. Senate runoff May 26: Cornyn 41.9% vs Paxton 40.7%. Crenshaw lost."},
{s:"UT",n:"Utah",filing:"Jan 8 / Mar 13, 2026",primary:"Jun 23, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:false,govOpen:false,govNote:"",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Legislature"],
centralDb:{name:"UT Lt. Governor's Office",url:"https://vote.utah.gov/candidates/",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Clerk",localCount:29,localDir:"https://vote.utah.gov/county-clerk-contact-information/",
counties:[["Beaver","https://vote.utah.gov/county-clerk-contact-information/"],["Box Elder","https://vote.utah.gov/county-clerk-contact-information/"],["Cache","https://vote.utah.gov/county-clerk-contact-information/"],["Carbon","https://vote.utah.gov/county-clerk-contact-information/"],["Daggett","https://vote.utah.gov/county-clerk-contact-information/"],["Davis","https://vote.utah.gov/county-clerk-contact-information/"],["Duchesne","https://vote.utah.gov/county-clerk-contact-information/"],["Emery","https://vote.utah.gov/county-clerk-contact-information/"],["Garfield","https://vote.utah.gov/county-clerk-contact-information/"],["Grand","https://vote.utah.gov/county-clerk-contact-information/"],["Iron","https://vote.utah.gov/county-clerk-contact-information/"],["Juab","https://vote.utah.gov/county-clerk-contact-information/"],["Kane","https://vote.utah.gov/county-clerk-contact-information/"],["Millard","https://vote.utah.gov/county-clerk-contact-information/"],["Morgan","https://vote.utah.gov/county-clerk-contact-information/"],["Piute","https://vote.utah.gov/county-clerk-contact-information/"],["Rich","https://vote.utah.gov/county-clerk-contact-information/"],["Salt Lake","https://vote.utah.gov/county-clerk-contact-information/"],["San Juan","https://vote.utah.gov/county-clerk-contact-information/"],["Sanpete","https://vote.utah.gov/county-clerk-contact-information/"],["Sevier","https://vote.utah.gov/county-clerk-contact-information/"],["Summit","https://vote.utah.gov/county-clerk-contact-information/"],["Tooele","https://vote.utah.gov/county-clerk-contact-information/"],["Uintah","https://vote.utah.gov/county-clerk-contact-information/"],["Utah","https://vote.utah.gov/county-clerk-contact-information/"],["Wasatch","https://vote.utah.gov/county-clerk-contact-information/"],["Washington","https://vote.utah.gov/county-clerk-contact-information/"],["Wayne","https://vote.utah.gov/county-clerk-contact-information/"],["Weber","https://vote.utah.gov/county-clerk-contact-information/"]],
gap:"County, municipal, school board at county clerk level. Only 29 counties.",
schoolNote:"Data held by county clerks.",
notes:"No Gov or Senate. Dual-path: signatures by Apr 11 OR convention Apr 25. Only 29 counties."},
{s:"VT",n:"Vermont",filing:"May 28, 2026",primary:"Aug 11, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:false,govNote:"Scott (R) 2-yr term",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Auditor","Full Legislature"],
centralDb:{name:"VT Secretary of State",url:"https://sos.vermont.gov/elections/candidates/",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"Town Clerk",localCount:246,localDir:"https://sos.vermont.gov/elections/town-clerks/",
counties:[["Addison","https://sos.vermont.gov/elections/town-clerks/"],["Bennington","https://sos.vermont.gov/elections/town-clerks/"],["Caledonia","https://sos.vermont.gov/elections/town-clerks/"],["Chittenden","https://sos.vermont.gov/elections/town-clerks/"],["Essex","https://sos.vermont.gov/elections/town-clerks/"],["Franklin","https://sos.vermont.gov/elections/town-clerks/"],["Grand Isle","https://sos.vermont.gov/elections/town-clerks/"],["Lamoille","https://sos.vermont.gov/elections/town-clerks/"],["Orange","https://sos.vermont.gov/elections/town-clerks/"],["Orleans","https://sos.vermont.gov/elections/town-clerks/"],["Rutland","https://sos.vermont.gov/elections/town-clerks/"],["Washington","https://sos.vermont.gov/elections/town-clerks/"],["Windham","https://sos.vermont.gov/elections/town-clerks/"],["Windsor","https://sos.vermont.gov/elections/town-clerks/"]],
gap:"All local and school board at town level. 246 towns.",
schoolNote:"Data held by supervisory union or town clerks. Town Meeting Day tradition.",
notes:"Governor every 2 years. Scott (R) in deep blue state. 246 towns."},
{s:"VA",n:"Virginia",filing:"Apr 2 / May 26, 2026",primary:"Aug 4, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:false,govOpen:false,govNote:"",senate:true,senOpen:false,senNote:"Warner (D)",house:true,
otherOffices:["No legislative races (off-cycle)"],
centralDb:{name:"VA State Board of Elections",url:"https://www.elections.virginia.gov/candidatepac-info/",covers:"Federal, statewide",format:"Online listings + some downloadable data",bulk:true},
localType:"General Registrar (County/City)",localCount:133,localDir:"https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup",
counties:[["Accomack","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Albemarle","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Alleghany","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Amelia","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Amherst","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Appomattox","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Arlington","https://vote.arlingtonva.us/"],["Augusta","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Bath","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Bedford","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Bland","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Botetourt","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Bristol City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Brunswick","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Buchanan","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Buckingham","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Campbell","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Caroline","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Carroll","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Charles City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Charlotte","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Charlottesville City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Chesapeake City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Chesterfield","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Clarke","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Craig","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Culpeper","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Cumberland","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Danville City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Dickenson","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Dinwiddie","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Essex","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Fairfax","https://www.fairfaxcounty.gov/elections/"],["Fairfax City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Falls Church City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Fauquier","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Floyd","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Fluvanna","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Franklin","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Franklin City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Frederick","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Fredericksburg City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Galax City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Giles","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Gloucester","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Goochland","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Grayson","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Greene","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Greensville","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Halifax","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Hampton City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Hanover","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Harrisonburg City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Henrico","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Henry","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Highland","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Hopewell City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Isle of Wight","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["James City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["King George","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["King William","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["King and Queen","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Lancaster","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Lee","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Lexington City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Loudoun","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Louisa","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Lunenburg","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Lynchburg City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Madison","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Manassas City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Manassas Park City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Martinsville City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Mathews","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Mecklenburg","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Middlesex","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Montgomery","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Nelson","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["New Kent","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Newport News City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Norfolk City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Northampton","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Northumberland","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Norton City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Nottoway","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Orange","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Page","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Patrick","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Petersburg City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Pittsylvania","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Poquoson City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Portsmouth City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Powhatan","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Prince Edward","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Prince George","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Prince William","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Pulaski","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Radford City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Rappahannock","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Richmond","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Richmond City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Roanoke","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Roanoke City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Rockbridge","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Rockingham","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Russell","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Salem City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Scott","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Shenandoah","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Smyth","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Southampton","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Spotsylvania","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Stafford","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Staunton City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Suffolk City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Surry","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Sussex","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Tazewell","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Virginia Beach City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Warren","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Washington","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Waynesboro City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Westmoreland","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Williamsburg City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Winchester City","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Wise","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["Wythe","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"],["York","https://vote.elections.virginia.gov/VoterInformation/PublicContactLookup"]],
gap:"County, municipal, school board at general registrar level. 133 jurisdictions.",
schoolNote:"Some localities elect, some appoint. Elected: data held by general registrar.",
notes:"No Governor (odd-year). No legislature. Parties choose convention or primary. 133 jurisdictions."},
{s:"WA",n:"Washington",filing:"May 8, 2026",primary:"Aug 4, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:false,govOpen:false,govNote:"",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Legislature"],
centralDb:{name:"VoteWA Candidate List",url:"https://voter.votewa.gov/CandidateList.aspx",covers:"★ ALL levels: federal, state, county, municipal, school board, fire district, special district",format:"Online searchable list + address-based ballot lookup",bulk:true},
localType:"N/A — Fully centralized",localCount:0,localDir:"https://voter.votewa.gov/CandidateList.aspx",
gap:"NONE — Washington centralizes ALL candidate data through VoteWA. County auditors feed into the state system.",
schoolNote:"★ School board candidates ARE in the VoteWA system.",
notes:"★ EXCELLENT centralized data via VoteWA. No Gov or Senate. Top-two primary. All-mail voting."},
{s:"WV",n:"West Virginia",filing:"Jan 31, 2026",primary:"May 12, 2026",runoff:null,general:"Nov 3, 2026",status:"closed",
gov:false,govOpen:false,govNote:"",senate:true,senOpen:false,senNote:"Capito (R)",house:true,
otherOffices:["Legislature"],
centralDb:{name:"WV Secretary of State",url:"https://apps.sos.wv.gov/elections/candidate-search/",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Clerk",localCount:55,localDir:"https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx",
counties:[["Barbour","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Berkeley","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Boone","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Braxton","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Brooke","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Cabell","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Calhoun","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Clay","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Doddridge","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Fayette","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Gilmer","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Grant","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Greenbrier","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Hampshire","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Hancock","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Hardy","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Harrison","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Jackson","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Jefferson","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Kanawha","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Lewis","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Lincoln","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Logan","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Marion","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Marshall","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Mason","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["McDowell","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Mercer","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Mineral","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Mingo","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Monongalia","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Monroe","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Morgan","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Nicholas","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Ohio","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Pendleton","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Pleasants","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Pocahontas","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Preston","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Putnam","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Raleigh","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Randolph","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Ritchie","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Roane","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Summers","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Taylor","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Tucker","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Tyler","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Upshur","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Wayne","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Webster","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Wetzel","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Wirt","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Wood","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"],["Wyoming","https://sos.wv.gov/elections/Pages/CountyClerkDirectory.aspx"]],
gap:"County, municipal, school board at county clerk level.",
schoolNote:"Data held by county clerks. Nonpartisan races.",
notes:"No Governor. Senate and House plus legislature."},
{s:"WI",n:"Wisconsin",filing:"Jun 1, 2026",primary:"Aug 11, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:true,govNote:"Evers not running",senate:false,senOpen:false,senNote:"No Senate race",house:true,
otherOffices:["Attorney General","Secretary of State","Treasurer","Legislature"],
centralDb:{name:"WI Elections Commission",url:"https://elections.wi.gov/candidates",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Clerk / Municipal Clerk",localCount:72,localDir:"https://elections.wi.gov/clerks/directory",
counties:[["Adams","https://elections.wi.gov/clerks/directory"],["Ashland","https://elections.wi.gov/clerks/directory"],["Barron","https://elections.wi.gov/clerks/directory"],["Bayfield","https://elections.wi.gov/clerks/directory"],["Brown","https://elections.wi.gov/clerks/directory"],["Buffalo","https://elections.wi.gov/clerks/directory"],["Burnett","https://elections.wi.gov/clerks/directory"],["Calumet","https://elections.wi.gov/clerks/directory"],["Chippewa","https://elections.wi.gov/clerks/directory"],["Clark","https://elections.wi.gov/clerks/directory"],["Columbia","https://elections.wi.gov/clerks/directory"],["Crawford","https://elections.wi.gov/clerks/directory"],["Dane","https://elections.wi.gov/clerks/directory"],["Dodge","https://elections.wi.gov/clerks/directory"],["Door","https://elections.wi.gov/clerks/directory"],["Douglas","https://elections.wi.gov/clerks/directory"],["Dunn","https://elections.wi.gov/clerks/directory"],["Eau Claire","https://elections.wi.gov/clerks/directory"],["Florence","https://elections.wi.gov/clerks/directory"],["Fond du Lac","https://elections.wi.gov/clerks/directory"],["Forest","https://elections.wi.gov/clerks/directory"],["Grant","https://elections.wi.gov/clerks/directory"],["Green","https://elections.wi.gov/clerks/directory"],["Green Lake","https://elections.wi.gov/clerks/directory"],["Iowa","https://elections.wi.gov/clerks/directory"],["Iron","https://elections.wi.gov/clerks/directory"],["Jackson","https://elections.wi.gov/clerks/directory"],["Jefferson","https://elections.wi.gov/clerks/directory"],["Juneau","https://elections.wi.gov/clerks/directory"],["Kenosha","https://elections.wi.gov/clerks/directory"],["Kewaunee","https://elections.wi.gov/clerks/directory"],["La Crosse","https://elections.wi.gov/clerks/directory"],["Lafayette","https://elections.wi.gov/clerks/directory"],["Langlade","https://elections.wi.gov/clerks/directory"],["Lincoln","https://elections.wi.gov/clerks/directory"],["Manitowoc","https://elections.wi.gov/clerks/directory"],["Marathon","https://elections.wi.gov/clerks/directory"],["Marinette","https://elections.wi.gov/clerks/directory"],["Marquette","https://elections.wi.gov/clerks/directory"],["Menominee","https://elections.wi.gov/clerks/directory"],["Milwaukee","https://elections.wi.gov/clerks/directory"],["Monroe","https://elections.wi.gov/clerks/directory"],["Oconto","https://elections.wi.gov/clerks/directory"],["Oneida","https://elections.wi.gov/clerks/directory"],["Outagamie","https://elections.wi.gov/clerks/directory"],["Ozaukee","https://elections.wi.gov/clerks/directory"],["Pepin","https://elections.wi.gov/clerks/directory"],["Pierce","https://elections.wi.gov/clerks/directory"],["Polk","https://elections.wi.gov/clerks/directory"],["Portage","https://elections.wi.gov/clerks/directory"],["Price","https://elections.wi.gov/clerks/directory"],["Racine","https://elections.wi.gov/clerks/directory"],["Richland","https://elections.wi.gov/clerks/directory"],["Rock","https://elections.wi.gov/clerks/directory"],["Rusk","https://elections.wi.gov/clerks/directory"],["Sauk","https://elections.wi.gov/clerks/directory"],["Sawyer","https://elections.wi.gov/clerks/directory"],["Shawano","https://elections.wi.gov/clerks/directory"],["Sheboygan","https://elections.wi.gov/clerks/directory"],["St. Croix","https://elections.wi.gov/clerks/directory"],["Taylor","https://elections.wi.gov/clerks/directory"],["Trempealeau","https://elections.wi.gov/clerks/directory"],["Vernon","https://elections.wi.gov/clerks/directory"],["Vilas","https://elections.wi.gov/clerks/directory"],["Walworth","https://elections.wi.gov/clerks/directory"],["Washburn","https://elections.wi.gov/clerks/directory"],["Washington","https://elections.wi.gov/clerks/directory"],["Waukesha","https://elections.wi.gov/clerks/directory"],["Waupaca","https://elections.wi.gov/clerks/directory"],["Waushara","https://elections.wi.gov/clerks/directory"],["Winnebago","https://elections.wi.gov/clerks/directory"],["Wood","https://elections.wi.gov/clerks/directory"]],
gap:"County, municipal, school board at county/municipal clerk level. School boards April elections.",
schoolNote:"Data held by school district clerks. April elections (separate cycle).",
notes:"Evers (D) not running. Open governor's race in key swing state."},
{s:"WY",n:"Wyoming",filing:"May 29, 2026",primary:"Aug 18, 2026",runoff:null,general:"Nov 3, 2026",status:"open",
gov:true,govOpen:true,govNote:"Gordon term-limited",senate:true,senOpen:true,senNote:"Open — Lummis retiring",house:true,
otherOffices:["Secretary of State","Auditor","Treasurer","Supt. of Public Instruction","Legislature"],
centralDb:{name:"WY Secretary of State",url:"https://sos.wyo.gov/elections/",covers:"Federal, statewide, legislature",format:"Online listings",bulk:false},
localType:"County Clerk",localCount:23,localDir:"https://sos.wyo.gov/elections/docs/wycountyclerks.pdf",
counties:[["Albany","https://sos.wyo.gov/elections/"],["Big Horn","https://sos.wyo.gov/elections/"],["Campbell","https://sos.wyo.gov/elections/"],["Carbon","https://sos.wyo.gov/elections/"],["Converse","https://sos.wyo.gov/elections/"],["Crook","https://sos.wyo.gov/elections/"],["Fremont","https://sos.wyo.gov/elections/"],["Goshen","https://sos.wyo.gov/elections/"],["Hot Springs","https://sos.wyo.gov/elections/"],["Johnson","https://sos.wyo.gov/elections/"],["Laramie","https://sos.wyo.gov/elections/"],["Lincoln","https://sos.wyo.gov/elections/"],["Natrona","https://sos.wyo.gov/elections/"],["Niobrara","https://sos.wyo.gov/elections/"],["Park","https://sos.wyo.gov/elections/"],["Platte","https://sos.wyo.gov/elections/"],["Sheridan","https://sos.wyo.gov/elections/"],["Sublette","https://sos.wyo.gov/elections/"],["Sweetwater","https://sos.wyo.gov/elections/"],["Teton","https://sos.wyo.gov/elections/"],["Uinta","https://sos.wyo.gov/elections/"],["Washakie","https://sos.wyo.gov/elections/"],["Weston","https://sos.wyo.gov/elections/"]],
gap:"Only 23 counties — trivial to cover manually.",
schoolNote:"Data held by county clerks.",
notes:"Gordon term-limited. Lummis retiring. Both seats open. Only 23 counties."},
];

function isWithinDays(dateStr,days){
  if(!dateStr) return false;
  const parts=dateStr.split("/");
  const now=new Date();
  const cutoff=new Date(now.getTime()+days*24*60*60*1000);
  return parts.some(p=>{
    let s=p.trim();
    if(!/\d{4}/.test(s)){
      const yearMatch=dateStr.match(/\d{4}/);
      if(yearMatch) s+=", "+yearMatch[0];
    }
    const d=new Date(s);
    return !isNaN(d)&&d>=now&&d<=cutoff;
  });
}
function hasElectionWithin(d,days){
  return isWithinDays(d.primary,days)||isWithinDays(d.general,days)||isWithinDays(d.runoff,days);
}

const STATUS = {
  voted:{l:"Primary Held",c:"#059669",bg:"#ecfdf5",i:"✓"},
  closed:{l:"Filing Closed",c:"#d97706",bg:"#fffbeb",i:"◉"},
  open:{l:"Filing Open",c:"#2563eb",bg:"#eff6ff",i:"○"},
  partial:{l:"Partially Closed",c:"#7c3aed",bg:"#f5f3ff",i:"◐"},
};


function App(){
  const [authed,setAuthed]=useState(()=>typeof sessionStorage!=="undefined"&&sessionStorage.getItem("eh_auth")==="1");
  const [loginPw,setLoginPw]=useState("");
  const [loginErr,setLoginErr]=useState("");

  // Hash-based routing for browser back/forward
  function parseHash(){
    const h=window.location.hash.slice(1);
    if(h.startsWith("state/")){return{view:"state",sel:h.slice(6).toUpperCase()||"NC"};}
    return{view:"dashboard",sel:"NC"};
  }
  const initial=typeof window!=="undefined"&&window.location.hash?parseHash():{view:"dashboard",sel:"NC"};
  const [view,setViewRaw]=useState(initial.view);
  const [sel,setSelRaw]=useState(initial.sel);

  const navigate=(newView,newSel)=>{
    const hash=newView==="state"?`#state/${newSel||sel}`:"#dashboard";
    window.history.pushState(null,"",hash);
    setViewRaw(newView);
    if(newSel) setSelRaw(newSel);
  };
  const setView=(v)=>navigate(v);
  const setSel=(s)=>{setSelRaw(s);navigate("state",s);};

  useEffect(()=>{
    const onPop=()=>{const p=parseHash();setViewRaw(p.view);setSelRaw(p.sel);};
    window.addEventListener("popstate",onPop);
    return()=>window.removeEventListener("popstate",onPop);
  },[]);
  const [search,setSearch]=useState("");
  const [filters,setFilters]=useState(new Set());
  const [mobile,setMobile]=useState(()=>typeof window!=="undefined"&&window.innerWidth<768);

  useEffect(()=>{
    const onResize=()=>setMobile(window.innerWidth<768);
    window.addEventListener("resize",onResize);
    return ()=>window.removeEventListener("resize",onResize);
  },[]);

  const [candidates,setCandidates]=useState([]);
  const [candLoading,setCandLoading]=useState(false);
  const [candFilter,setCandFilter]=useState("");
  const [candLevelFilter,setCandLevelFilter]=useState("all");
  const [sortCol,setSortCol]=useState("name");
  const [sortAsc,setSortAsc]=useState(true);

  const [uploadStatus,setUploadStatus]=useState(null);

  // Load candidate data from API when state changes (automated + manual uploads)
  const loadCandidates=useCallback((state)=>{
    if(!state) return;
    setCandidates([]);
    setCandLoading(true);
    const stateCollectors={NC:"/api/collect/nc",MN:"/api/collect/mn",CA:"/api/collect/ca",GA:"/api/collect/ga",MD:"/api/collect/md",OR:"/api/collect/or"};
    const fetches=[
      fetch(`/api/collect/fec?state=${state}`).then(r=>r.ok?r.json():null).catch(()=>null),
      // Always check for manual uploads
      fetch(`/api/history?source=manual_${state.toLowerCase()}&state=${state}`).then(r=>r.ok?r.json():null).catch(()=>null),
    ];
    if(stateCollectors[state]){
      fetches.push(fetch(stateCollectors[state]).then(r=>r.ok?r.json():null).catch(()=>null));
    }
    Promise.all(fetches).then(async results=>{
      const byKey=new Map();
      for(const result of results){
        if(!result) continue;
        // History endpoint returns metadata, not data — fetch the latest snapshot
        if(result.latest&&result.latest.date){
          const snap=await fetch(`/api/history?source=manual_${state.toLowerCase()}&state=${state}&date=${result.latest.date}`).then(r=>r.ok?r.json():null).catch(()=>null);
          if(snap&&Array.isArray(snap.data)){
            for(const c of snap.data) byKey.set(`${c.name}|${c.office}`,c);
          }
          continue;
        }
        const list=Array.isArray(result.data)?result.data:[];
        for(const c of list) byKey.set(`${c.name}|${c.office}`,c);
      }
      setCandidates([...byKey.values()]);
      setCandLoading(false);
    });
  },[]);

  const handleUpload=useCallback(async(file)=>{
    if(!file||!sel) return;
    setUploadStatus("Uploading...");
    const text=await file.text();
    try{
      const resp=await fetch("/api/upload",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({state:sel,csv:text}),
      });
      const result=await resp.json();
      if(result.success){
        setUploadStatus(`Uploaded ${result.uploaded} rows: ${result.newCandidates} new, ${result.enrichedExisting} enriched. Total: ${result.totalAfterMerge}`);
        loadCandidates(sel);
      } else {
        setUploadStatus(`Error: ${result.error}`);
      }
    }catch(err){
      setUploadStatus(`Upload failed: ${err.message}`);
    }
  },[sel,loadCandidates]);

  useEffect(()=>{ loadCandidates(sel); },[sel,loadCandidates]);

  const toggleSort=(col)=>{
    if(sortCol===col) setSortAsc(!sortAsc);
    else { setSortCol(col); setSortAsc(true); }
  };

  const filteredCandidates=useMemo(()=>{
    const list=candidates.filter(c=>{
      if(candLevelFilter!=="all"&&c.officeLevel!==candLevelFilter) return false;
      if(candFilter){
        const q=candFilter.toLowerCase();
        return c.name.toLowerCase().includes(q)||c.office.toLowerCase().includes(q)||(c.party||"").toLowerCase().includes(q);
      }
      return true;
    });
    list.sort((a,b)=>{
      let va=a[sortCol]||"", vb=b[sortCol]||"";
      if(typeof va==="string") va=va.toLowerCase();
      if(typeof vb==="string") vb=vb.toLowerCase();
      if(va<vb) return sortAsc?-1:1;
      if(va>vb) return sortAsc?1:-1;
      return 0;
    });
    return list;
  },[candidates,candFilter,candLevelFilter,sortCol,sortAsc]);

  const exportExcel=useCallback(()=>{
    const rows=filteredCandidates.map(c=>({
      Name:c.name,"First Name":c.firstName||"","Last Name":c.lastName||"",
      Office:c.office,Level:c.officeLevel||"",Party:c.party||"",County:c.county||"",
      Email:c.email||"",Phone:c.phone||"",Website:c.website||"",
      Address:c.address||"",City:c.city||"",State:c.state||sel,Zip:c.zip||"",
      Source:c.source||"",
    }));
    const ws=XLSX.utils.json_to_sheet(rows);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,`${sel} Candidates`);
    XLSX.writeFile(wb,`${sel}_candidates_2026.xlsx`);
  },[filteredCandidates,sel]);

  const toggleFilter=(k)=>setFilters(prev=>{
    const next=new Set(prev);
    if(next.has(k)) next.delete(k); else next.add(k);
    return next;
  });

  const filtered=useMemo(()=>S.filter(d=>{
    const m1=!search||d.n.toLowerCase().includes(search.toLowerCase())||d.s.toLowerCase().includes(search.toLowerCase());
    if(filters.size===0) return m1;
    let m2=false;
    if(filters.has("voted")&&d.status==="voted") m2=true;
    if(filters.has("closed")&&d.status==="closed") m2=true;
    if(filters.has("open")&&(d.status==="open"||d.status==="partial")) m2=true;
    if(filters.has("star")&&["NC","IL","MN","LA","WA"].includes(d.s)) m2=true;
    if(filters.has("in30")&&hasElectionWithin(d,30)) m2=true;
    if(filters.has("in45")&&hasElectionWithin(d,45)) m2=true;
    if(filters.has("in60")&&hasElectionWithin(d,60)) m2=true;
    if(filters.has("in90")&&hasElectionWithin(d,90)) m2=true;
    return m1&&m2;
  }),[search,filters]);

  const cur=S.find(d=>d.s===sel);
  const st=cur?STATUS[cur.status]:null;
  const isStar=cur&&["NC","IL","MN","LA","WA"].includes(cur.s);

  const handleLogin=async(e)=>{
    e.preventDefault();
    setLoginErr("");
    try{
      const r=await fetch("/api/auth",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({password:loginPw})});
      if(r.ok){setAuthed(true);sessionStorage.setItem("eh_auth","1");}
      else setLoginErr("Invalid password");
    }catch{setLoginErr("Login failed");}
  };

  const selectState=(code)=>setSel(code);

  // Compute dashboard stats
  const battlegrounds=S.filter(d=>["GA","AZ","MI","WI","PA","NV","NC","NH","ME","OH"].includes(d.s));
  const openGov=S.filter(d=>d.govOpen);
  const openSenate=S.filter(d=>d.senOpen);
  const soonStates=S.filter(d=>hasElectionWithin(d,60));

  // --- Auth gate ---
  if(!authed) return(
    <div style={{fontFamily:"'DM Sans',system-ui,sans-serif",minHeight:"100vh",background:"linear-gradient(135deg,#0c1222 0%,#1a1a3e 100%)",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <form onSubmit={handleLogin} style={{background:"#fff",borderRadius:16,padding:"40px 36px",width:360,maxWidth:"90vw",boxShadow:"0 20px 60px rgba(0,0,0,0.3)"}}>
        <h1 style={{fontFamily:"'Newsreader',serif",fontSize:22,fontWeight:800,color:"#0f172a",margin:"0 0 4px",textAlign:"center"}}>2026 Election Research Hub</h1>
        <p style={{fontSize:12,color:"#64748b",textAlign:"center",margin:"0 0 24px"}}>PolitiFast Internal Tool</p>
        <input type="password" value={loginPw} onChange={e=>setLoginPw(e.target.value)} placeholder="Enter password"
          style={{width:"100%",padding:"12px 16px",borderRadius:8,border:"1px solid #d1d5db",fontSize:14,marginBottom:12,outline:"none"}} autoFocus/>
        {loginErr&&<div style={{color:"#dc2626",fontSize:12,marginBottom:8,textAlign:"center"}}>{loginErr}</div>}
        <button type="submit" style={{width:"100%",padding:"12px",borderRadius:8,background:"#f59e0b",color:"#fff",fontWeight:700,fontSize:14,border:"none",cursor:"pointer"}}>Sign In</button>
      </form>
    </div>
  );

  return(
  <div style={{fontFamily:"'DM Sans',system-ui,sans-serif",minHeight:"100vh",background:"#f5f6f8"}}>


  <header style={{background:"linear-gradient(135deg,#0c1222 0%,#1a1a3e 100%)",padding:mobile?"14px 12px 10px":"18px 24px 14px",borderBottom:"3px solid #f59e0b"}}>
    <div style={{maxWidth:1300,margin:"0 auto"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
        <div>
          <h1 onClick={()=>setView("dashboard")} style={{fontFamily:"'Newsreader',serif",fontSize:mobile?19:24,fontWeight:800,color:"#fff",margin:"0 0 2px",letterSpacing:-.3,cursor:"pointer"}}>2026 Election Research Hub</h1>
          <p style={{fontSize:mobile?10:12,color:"#94a3b8",margin:0}}>{mobile?"All 50 states":"Filing deadlines · Primary dates · Candidate data sources · County office directories — all 50 states"}</p>
        </div>
        <div style={{display:"flex",gap:6}}>
          <button onClick={()=>setView("dashboard")} style={{padding:"5px 14px",borderRadius:6,fontSize:11,fontWeight:700,border:"none",cursor:"pointer",background:view==="dashboard"?"#f59e0b":"rgba(255,255,255,0.1)",color:view==="dashboard"?"#0c1222":"#94a3b8"}}>Dashboard</button>
          <button onClick={()=>setView("state")} style={{padding:"5px 14px",borderRadius:6,fontSize:11,fontWeight:700,border:"none",cursor:"pointer",background:view==="state"?"#f59e0b":"rgba(255,255,255,0.1)",color:view==="state"?"#0c1222":"#94a3b8"}}>State Detail</button>
          <button onClick={()=>{setAuthed(false);sessionStorage.removeItem("eh_auth");}} style={{padding:"5px 14px",borderRadius:6,fontSize:11,fontWeight:700,border:"none",cursor:"pointer",background:"rgba(255,255,255,0.08)",color:"#64748b"}}>Logout</button>
        </div>
      </div>
      <div style={{display:"flex",gap:mobile?10:16,marginTop:mobile?8:12,flexWrap:"wrap"}}>
        {[
          {n:S.filter(d=>d.status==="voted").length,l:"Primaries Held",c:"#4ade80"},
          {n:S.filter(d=>d.status==="closed").length,l:"Filing Closed",c:"#fbbf24"},
          {n:S.filter(d=>d.status==="open"||d.status==="partial").length,l:"Filing Open",c:"#60a5fa"},
          {n:5,l:"Full Data Centralization",c:"#c084fc"},
        ].map(({n,l,c})=>(
          <div key={l} style={{display:"flex",alignItems:"center",gap:6}}>
            <span style={{width:7,height:7,borderRadius:"50%",background:c}}/>
            <span style={{fontSize:18,fontWeight:800,color:"#fff",fontFamily:"'JetBrains Mono'"}}>{n}</span>
            <span style={{fontSize:11,color:"#94a3b8"}}>{l}</span>
          </div>
        ))}
      </div>
    </div>
  </header>

  {/* === DASHBOARD VIEW === */}
  {view==="dashboard"&&(
  <div style={{maxWidth:1300,margin:"0 auto",padding:mobile?"12px 8px":"20px 16px"}}>

    {/* Key Metrics Row */}
    <div style={{display:"grid",gridTemplateColumns:mobile?"repeat(2,1fr)":"repeat(4,1fr)",gap:10,marginBottom:16}}>
      {[
        {n:openGov.length,l:"Open Governor Races",c:"#dc2626",bg:"#fef2f2",sub:openGov.slice(0,5).map(d=>d.s).join(", ")+(openGov.length>5?`… +${openGov.length-5}`:"")},
        {n:openSenate.length,l:"Open Senate Seats",c:"#7c3aed",bg:"#f5f3ff",sub:openSenate.map(d=>d.s).join(", ")},
        {n:soonStates.length,l:"Elections <60 Days",c:"#d97706",bg:"#fffbeb",sub:soonStates.slice(0,5).map(d=>d.s).join(", ")+(soonStates.length>5?`… +${soonStates.length-5}`:"")},
        {n:battlegrounds.length,l:"Key Battlegrounds",c:"#1d4ed8",bg:"#eff6ff",sub:battlegrounds.map(d=>d.s).join(", ")},
      ].map(({n,l,c,bg,sub})=>(
        <div key={l} style={{background:bg,border:`1px solid ${c}22`,borderRadius:12,padding:"16px 18px"}}>
          <div style={{fontSize:28,fontWeight:800,color:c,fontFamily:"'JetBrains Mono'"}}>{n}</div>
          <div style={{fontSize:12,fontWeight:700,color:c,marginBottom:4}}>{l}</div>
          <div style={{fontSize:10,color:"#6b7280"}}>{sub}</div>
        </div>
      ))}
    </div>

    {/* Battleground States */}
    <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"16px 20px",marginBottom:16}}>
      <div style={{fontSize:11,fontWeight:700,color:"#1d4ed8",letterSpacing:1,textTransform:"uppercase",marginBottom:12}}>Key Battleground States</div>
      <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"repeat(2,1fr)",gap:8}}>
        {battlegrounds.map(d=>{
          const t=STATUS[d.status];
          return(
            <div key={d.s} onClick={()=>selectState(d.s)} style={{
              display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderRadius:8,
              background:"#f8fafc",border:"1px solid #e5e7eb",cursor:"pointer",
            }}>
              <span style={{fontFamily:"'JetBrains Mono'",fontSize:14,fontWeight:800,color:"#0f172a",minWidth:28}}>{d.s}</span>
              <div style={{flex:1}}>
                <div style={{fontSize:13,fontWeight:600,color:"#0f172a"}}>{d.n}</div>
                <div style={{fontSize:10,color:"#6b7280"}}>
                  {d.gov&&d.govOpen?"Gov (Open) · ":""}
                  {d.senate&&d.senOpen?"Senate (Open) · ":""}
                  {d.senate&&!d.senOpen&&d.senNote?`Senate: ${d.senNote} · `:""}
                  Primary: {d.primary}
                </div>
              </div>
              <span style={{padding:"2px 8px",borderRadius:10,fontSize:9,fontWeight:700,color:t.c,background:t.bg}}>{t.l}</span>
            </div>
          );
        })}
      </div>
    </div>

    {/* Open Seats Grid */}
    <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"1fr 1fr",gap:16,marginBottom:16}}>
      <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"16px 20px"}}>
        <div style={{fontSize:11,fontWeight:700,color:"#dc2626",letterSpacing:1,textTransform:"uppercase",marginBottom:10}}>Open Governor Races ({openGov.length})</div>
        {openGov.map(d=>(
          <div key={d.s} onClick={()=>selectState(d.s)} style={{
            display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid #f1f5f9",cursor:"pointer",fontSize:12,
          }}>
            <span><strong>{d.s}</strong> {d.n}</span>
            <span style={{color:"#6b7280",fontSize:11}}>{d.govNote||"Open"}</span>
          </div>
        ))}
      </div>
      <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"16px 20px"}}>
        <div style={{fontSize:11,fontWeight:700,color:"#7c3aed",letterSpacing:1,textTransform:"uppercase",marginBottom:10}}>Open Senate Seats ({openSenate.length})</div>
        {openSenate.map(d=>(
          <div key={d.s} onClick={()=>selectState(d.s)} style={{
            display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid #f1f5f9",cursor:"pointer",fontSize:12,
          }}>
            <span><strong>{d.s}</strong> {d.n}</span>
            <span style={{color:"#6b7280",fontSize:11}}>{d.senNote||"Open"}</span>
          </div>
        ))}
      </div>
    </div>

    {/* Upcoming Elections */}
    {soonStates.length>0&&(
    <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"16px 20px",marginBottom:16}}>
      <div style={{fontSize:11,fontWeight:700,color:"#d97706",letterSpacing:1,textTransform:"uppercase",marginBottom:10}}>Elections Within 60 Days</div>
      <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"repeat(3,1fr)",gap:8}}>
        {soonStates.map(d=>(
          <div key={d.s} onClick={()=>selectState(d.s)} style={{
            padding:"10px 14px",borderRadius:8,background:"#fffbeb",border:"1px solid #fde68a",cursor:"pointer",
          }}>
            <div style={{fontWeight:700,fontSize:13,color:"#92400e"}}>{d.s} — {d.n}</div>
            <div style={{fontSize:11,color:"#78350f",marginTop:2}}>
              {d.primary&&hasElectionWithin(d,60)&&isWithinDays(d.primary,60)?`Primary: ${d.primary}`:""}
              {d.runoff&&isWithinDays(d.runoff,60)?` · Runoff: ${d.runoff}`:""}
            </div>
          </div>
        ))}
      </div>
    </div>
    )}

    {/* All States Quick Grid */}
    <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"16px 20px"}}>
      <div style={{fontSize:11,fontWeight:700,color:"#374151",letterSpacing:1,textTransform:"uppercase",marginBottom:10}}>All 50 States</div>
      <div style={{display:"grid",gridTemplateColumns:mobile?"repeat(3,1fr)":"repeat(10,1fr)",gap:4}}>
        {S.map(d=>{
          const t=STATUS[d.status];
          return(
            <button key={d.s} onClick={()=>selectState(d.s)} style={{
              padding:"8px 4px",borderRadius:6,border:`1px solid ${t.c}44`,background:t.bg,
              cursor:"pointer",textAlign:"center",fontSize:12,fontWeight:700,color:t.c,
            }}>{d.s}</button>
          );
        })}
      </div>
    </div>
  </div>
  )}

  {/* === STATE DETAIL VIEW === */}
  {view==="state"&&(
  <div style={{maxWidth:1300,margin:"0 auto",padding:mobile?"8px 6px":"12px 10px",display:"flex",flexDirection:mobile?"column":"row",gap:mobile?8:12}}>
    {/* Sidebar / Mobile Nav */}
    <div style={{width:mobile?"100%":220,flexShrink:0,display:"flex",flexDirection:"column",gap:6}}>
      {mobile?(
        <>
          <div style={{display:"flex",gap:6}}>
            <select value={sel} onChange={e=>setSel(e.target.value)} style={{
              flex:1,padding:"8px 10px",borderRadius:8,border:"1px solid #d1d5db",fontSize:14,fontWeight:600,
              background:"#fff",color:"#111827",appearance:"auto",
            }}>
              {filtered.map(d=><option key={d.s} value={d.s}>{d.s} — {d.n}</option>)}
            </select>
          </div>
          <div style={{display:"flex",gap:3,flexWrap:"wrap"}}>
            <button onClick={()=>setFilters(new Set())} title="Show all 50 states" style={{
              padding:"4px 10px",borderRadius:6,fontSize:11,fontWeight:700,cursor:"pointer",
              border:filters.size===0?"2px solid #f59e0b":"1px solid #e5e7eb",
              background:filters.size===0?"#fffbeb":"#fff",
              color:filters.size===0?"#92400e":"#6b7280",
            }}>All</button>
            {[
              {k:"in30",l:"<30 Days",t:"Primary, general, or runoff within 30 days"},
              {k:"in45",l:"<45 Days",t:"Primary, general, or runoff within 45 days"},
              {k:"in60",l:"<60 Days",t:"Primary, general, or runoff within 60 days"},
              {k:"in90",l:"<90 Days",t:"Primary, general, or runoff within 90 days"},
              {k:"voted",l:"Voted",t:"States where the primary has already been held"},
              {k:"closed",l:"Closed",t:"Filing deadline has passed"},
              {k:"open",l:"Open",t:"Filing window is still open for candidates"},
              {k:"star",l:"★ Best",t:"States that centralize ALL candidate data including local and school board"},
            ].map(f=>(
              <button key={f.k} onClick={()=>toggleFilter(f.k)} title={f.t} style={{
                padding:"4px 10px",borderRadius:6,fontSize:11,fontWeight:700,cursor:"pointer",
                border:filters.has(f.k)?"2px solid #f59e0b":"1px solid #e5e7eb",
                background:filters.has(f.k)?"#fffbeb":"#fff",
                color:filters.has(f.k)?"#92400e":"#6b7280",
              }}>{f.l}</button>
            ))}
          </div>
        </>
      ):(
        <>
          <input type="text" placeholder="Search states..." value={search} onChange={e=>setSearch(e.target.value)}
            style={{width:"100%",padding:"8px 12px",borderRadius:8,border:"1px solid #d1d5db",fontSize:13,outline:"none"}}/>
          <div style={{display:"flex",gap:3,flexWrap:"wrap"}}>
            <button onClick={()=>setFilters(new Set())} title="Show all 50 states" style={{
              padding:"3px 9px",borderRadius:6,fontSize:10,fontWeight:700,cursor:"pointer",
              border:filters.size===0?"2px solid #f59e0b":"1px solid #e5e7eb",
              background:filters.size===0?"#fffbeb":"#fff",
              color:filters.size===0?"#92400e":"#6b7280",
            }}>All</button>
            {[
              {k:"in30",l:"<30 Days",t:"Primary, general, or runoff within 30 days"},
              {k:"in45",l:"<45 Days",t:"Primary, general, or runoff within 45 days"},
              {k:"in60",l:"<60 Days",t:"Primary, general, or runoff within 60 days"},
              {k:"in90",l:"<90 Days",t:"Primary, general, or runoff within 90 days"},
              {k:"voted",l:"Voted",t:"States where the primary has already been held"},
              {k:"closed",l:"Closed",t:"Filing deadline has passed"},
              {k:"open",l:"Open",t:"Filing window is still open for candidates"},
              {k:"star",l:"★ Best Data",t:"States that centralize ALL candidate data including local and school board"},
            ].map(f=>(
              <button key={f.k} onClick={()=>toggleFilter(f.k)} title={f.t} style={{
                padding:"3px 9px",borderRadius:6,fontSize:10,fontWeight:700,cursor:"pointer",
                border:filters.has(f.k)?"2px solid #f59e0b":"1px solid #e5e7eb",
                background:filters.has(f.k)?"#fffbeb":"#fff",
                color:filters.has(f.k)?"#92400e":"#6b7280",
              }}>{f.l}</button>
            ))}
          </div>
          <div style={{flex:1,overflowY:"auto",maxHeight:"calc(100vh - 230px)",background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:4}}>
            {filtered.map(d=>{
              const t=STATUS[d.status];
              const star=["NC","IL","MN","LA","WA"].includes(d.s);
              return(
                <button key={d.s} onClick={()=>setSel(d.s)} style={{
                  display:"flex",alignItems:"center",gap:6,padding:"6px 8px",width:"100%",textAlign:"left",
                  background:sel===d.s?"#fffbeb":"transparent",border:sel===d.s?"2px solid #f59e0b":"2px solid transparent",
                  borderRadius:6,cursor:"pointer",
                }}>
                  <span style={{fontSize:9,color:t.c,fontWeight:700,minWidth:12}}>{t.i}</span>
                  <span style={{fontFamily:"'JetBrains Mono'",fontSize:11,fontWeight:600,color:"#111827",minWidth:22}}>{d.s}</span>
                  <span style={{fontSize:11,color:"#4b5563",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.n}</span>
                  {star&&<span style={{fontSize:9,color:"#f59e0b"}}>★</span>}
                </button>
              );
            })}
          </div>
        </>
      )}
    </div>

    {/* Detail */}
    {cur&&(
    <div style={{flex:1,minWidth:0,animation:"fadeIn 0.15s"}}>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",gap:mobile?6:10,marginBottom:2,flexWrap:"wrap"}}>
        <h2 style={{fontFamily:"'Newsreader',serif",fontSize:mobile?21:26,fontWeight:800,color:"#0f172a",margin:0}}>{cur.n}</h2>
        <span style={{padding:"2px 10px",borderRadius:12,fontSize:10,fontWeight:700,color:st.c,background:st.bg}}>{st.l}</span>
        {isStar&&<span style={{padding:"2px 10px",borderRadius:12,fontSize:10,fontWeight:700,color:"#7c3aed",background:"#ede9fe"}}>★ Full Centralization</span>}
      </div>
      <p style={{fontSize:12,color:"#64748b",margin:"2px 0 10px",lineHeight:1.5}}>{cur.notes}</p>

      {/* Key Dates */}
      <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:"12px 16px",marginBottom:10}}>
        <div style={{fontSize:10,fontWeight:700,color:"#94a3b8",letterSpacing:1,textTransform:"uppercase",marginBottom:8}}>Key Dates</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:8}}>
          {[
            {l:"Filing Deadline",v:cur.filing},
            {l:"Primary",v:cur.primary},
            {l:"Runoff",v:cur.runoff||"—"},
            {l:"General",v:cur.general},
          ].map(({l,v})=>(
            <div key={l} style={{background:"#f8fafc",borderRadius:8,padding:"8px 12px"}}>
              <div style={{fontSize:9,fontWeight:700,color:"#94a3b8",letterSpacing:0.8,textTransform:"uppercase"}}>{l}</div>
              <div style={{fontSize:14,fontWeight:700,color:"#0f172a",fontFamily:"'JetBrains Mono'",marginTop:2}}>{v}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Races on Ballot */}
      <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:mobile?"10px 12px":"12px 16px",marginBottom:10}}>
        <div style={{fontSize:10,fontWeight:700,color:"#94a3b8",letterSpacing:1,textTransform:"uppercase",marginBottom:8}}>Races on Ballot</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:mobile?4:6,marginBottom:8}}>
          {cur.gov&&(
            <span style={{padding:"4px 12px",borderRadius:8,fontSize:12,fontWeight:600,background:cur.govOpen?"#fef2f2":"#f0fdf4",color:cur.govOpen?"#991b1b":"#166534",border:`1px solid ${cur.govOpen?"#fecaca":"#bbf7d0"}`}}>
              Governor {cur.govOpen?"(Open)":"(Incumbent)"} {cur.govNote&&<span style={{fontWeight:400,fontSize:11}}> — {cur.govNote}</span>}
            </span>
          )}
          {cur.senate&&(
            <span style={{padding:"4px 12px",borderRadius:8,fontSize:12,fontWeight:600,background:cur.senOpen?"#fef2f2":"#f0fdf4",color:cur.senOpen?"#991b1b":"#166534",border:`1px solid ${cur.senOpen?"#fecaca":"#bbf7d0"}`}}>
              U.S. Senate {cur.senOpen?"(Open)":""} {cur.senNote&&<span style={{fontWeight:400,fontSize:11}}> — {cur.senNote}</span>}
            </span>
          )}
          {cur.house&&<span style={{padding:"4px 12px",borderRadius:8,fontSize:12,fontWeight:600,background:"#eff6ff",color:"#1e40af",border:"1px solid #bfdbfe"}}>U.S. House — All Districts</span>}
        </div>
        {cur.otherOffices.length>0&&(
          <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
            {cur.otherOffices.map((o,i)=>(
              <span key={i} style={{padding:"3px 10px",borderRadius:6,fontSize:11,background:"#f1f5f9",color:"#475569",border:"1px solid #e2e8f0"}}>{o}</span>
            ))}
          </div>
        )}
      </div>

      {/* Where to Get Data */}
      <div style={{background:"#fff",border:"1px solid #e5e7eb",borderLeft:"4px solid #22d3ee",borderRadius:"0 10px 10px 0",padding:"14px 18px",marginBottom:10}}>
        <div style={{fontSize:10,fontWeight:700,color:"#0891b2",letterSpacing:1,textTransform:"uppercase",marginBottom:6}}>Where to Get Data</div>
        <div style={{fontSize:15,fontWeight:700,color:"#0f172a",marginBottom:4}}>{cur.centralDb.name}</div>
        <div style={{fontSize:12,color:"#4b5563",marginBottom:2}}><strong>Covers:</strong> {cur.centralDb.covers}</div>
        <div style={{fontSize:12,color:"#4b5563",marginBottom:6}}>
          <strong>Format:</strong> {cur.centralDb.format}
          {cur.centralDb.bulk&&<span style={{marginLeft:8,padding:"1px 8px",borderRadius:10,fontSize:10,fontWeight:700,color:"#7c3aed",background:"#ede9fe"}}>BULK DOWNLOAD</span>}
        </div>
        <a href={cur.centralDb.url} target="_blank" rel="noopener noreferrer" style={{fontSize:12,fontFamily:"'JetBrains Mono'",wordBreak:"break-all"}}>
          {cur.centralDb.url.replace("https://","").replace("http://","")}
        </a>
      </div>

      {/* Data Gaps */}
      <div style={{background:cur.gap.startsWith("NONE")?"#ecfdf5":"#fffbeb",border:`1px solid ${cur.gap.startsWith("NONE")?"#a7f3d0":"#fde68a"}`,borderRadius:10,padding:"10px 14px",marginBottom:10}}>
        <div style={{fontSize:10,fontWeight:700,color:cur.gap.startsWith("NONE")?"#059669":"#b45309",letterSpacing:1,textTransform:"uppercase",marginBottom:3}}>
          {cur.gap.startsWith("NONE")?"✓ Complete Coverage":"⚠ What's NOT in the state system — requires county offices"}
        </div>
        <div style={{fontSize:12,color:"#374151",lineHeight:1.5}}>{cur.gap}</div>
      </div>

      {/* School Board */}
      <div style={{background:"#fff",border:"1px solid #e5e7eb",borderLeft:"4px solid #dc2626",borderRadius:"0 10px 10px 0",padding:"12px 16px",marginBottom:10}}>
        <div style={{fontSize:10,fontWeight:700,color:"#dc2626",letterSpacing:1,textTransform:"uppercase",marginBottom:3}}>School Board Data</div>
        <div style={{fontSize:12,color:"#374151",lineHeight:1.5}}>{cur.schoolNote}</div>
      </div>

      {/* FEC */}
      <div style={{background:"#eff6ff",border:"1px solid #bfdbfe",borderRadius:10,padding:"10px 14px",marginBottom:10}}>
        <div style={{fontSize:10,fontWeight:700,color:"#1d4ed8",letterSpacing:1,marginBottom:3}}>Federal Data (All States)</div>
        <div style={{fontSize:12,color:"#374151"}}>
          FEC bulk downloads: <a href="https://www.fec.gov/data/browse-data/" target="_blank" rel="noopener noreferrer">fec.gov/data/browse-data</a> · API: <a href="https://api.open.fec.gov/developers/" target="_blank" rel="noopener noreferrer">api.open.fec.gov</a> · Free, updated nightly.
        </div>
      </div>

      {/* County Offices */}
      {cur.localCount===0?(
        <div style={{background:"#ecfdf5",border:"1px solid #a7f3d0",borderRadius:10,padding:"20px",textAlign:"center"}}>
          <div style={{fontSize:20,marginBottom:6}}>✓</div>
          <div style={{fontSize:14,fontWeight:700,color:"#059669"}}>No county-level lookups needed!</div>
          <div style={{fontSize:12,color:"#374151",marginTop:4}}>This state centralizes all candidate data — including local and school board — in one system.</div>
          <a href={cur.centralDb.url} target="_blank" rel="noopener noreferrer" style={{display:"inline-block",marginTop:10,padding:"8px 20px",borderRadius:8,background:"#059669",color:"#fff",fontWeight:700,fontSize:13,textDecoration:"none"}}>
            Go to {cur.centralDb.name} →
          </a>
        </div>
      ):(
        <>
          <div style={{background:"#fff",border:"1px solid #e5e7eb",borderLeft:"4px solid #f59e0b",borderRadius:"0 10px 10px 0",padding:"14px 18px",marginBottom:10}}>
            <div style={{fontSize:10,fontWeight:700,color:"#b45309",letterSpacing:1,textTransform:"uppercase",marginBottom:6}}>
              County Offices — {cur.localCount} {cur.localType} Offices
            </div>

            {cur.counties&&cur.counties.length>0?(
              <div style={{display:"grid",gridTemplateColumns:mobile?"repeat(auto-fill,minmax(140px,1fr))":"repeat(auto-fill,minmax(200px,1fr))",gap:4,marginTop:8}}>
                {cur.counties.map(([name,url])=>(
                  <a key={name} href={url} target="_blank" rel="noopener noreferrer" style={{
                    display:"block",padding:"5px 10px",borderRadius:6,fontSize:12,color:"#92400e",
                    background:"#fffbeb",border:"1px solid #fde68a",textDecoration:"none",
                    overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",
                  }}>
                    {name}
                  </a>
                ))}
              </div>
            ):(
              <div style={{background:"#fef3c7",borderRadius:8,padding:"12px 16px"}}>
                <div style={{fontSize:12,color:"#78350f",marginBottom:6}}>
                  State-maintained directory of all {cur.localCount} {cur.localType.toLowerCase()} offices for local candidate data.
                </div>
                <a href={cur.localDir} target="_blank" rel="noopener noreferrer" style={{
                  display:"inline-block",padding:"8px 20px",borderRadius:8,background:"#92400e",color:"#fff",
                  fontWeight:700,fontSize:13,textDecoration:"none",fontFamily:"'JetBrains Mono'"
                }}>
                  Open Directory of {cur.localCount} Offices →
                </a>
              </div>
            )}
          </div>

          <div style={{background:"#f8fafc",border:"1px solid #e5e7eb",borderRadius:10,padding:"12px 16px"}}>
            <div style={{fontSize:10,fontWeight:700,color:"#64748b",letterSpacing:1,textTransform:"uppercase",marginBottom:6}}>Difficulty Assessment</div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              <span style={{padding:"4px 12px",borderRadius:8,fontSize:12,fontWeight:600,
                background:cur.localCount<=25?"#ecfdf5":cur.localCount<=75?"#fffbeb":"#fef2f2",
                color:cur.localCount<=25?"#059669":cur.localCount<=75?"#d97706":"#dc2626",
                border:`1px solid ${cur.localCount<=25?"#a7f3d0":cur.localCount<=75?"#fde68a":"#fecaca"}`
              }}>
                {cur.localCount<=25?"Easy":cur.localCount<=75?"Moderate":cur.localCount<=150?"Labor-intensive":"Very difficult"} — {cur.localCount} offices
              </span>
              {cur.localCount>100&&<span style={{padding:"4px 12px",borderRadius:8,fontSize:12,background:"#ede9fe",color:"#5b21b6",border:"1px solid #ddd6fe"}}>Consider BallotReady or Ballotpedia subscription</span>}
            </div>
          </div>
        </>
      )}
      {/* Candidate Data */}
      <div style={{background:"#fff",border:"1px solid #e5e7eb",borderLeft:"4px solid #1d4ed8",borderRadius:"0 10px 10px 0",padding:"14px 18px",marginTop:10}}>
        <div style={{display:"flex",alignItems:"center",flexWrap:"wrap",gap:8,marginBottom:10}}>
          <div style={{fontSize:10,fontWeight:700,color:"#1d4ed8",letterSpacing:1,textTransform:"uppercase",marginRight:"auto"}}>
            Candidate Data {candidates.length>0?`— ${filteredCandidates.length} candidates`:""}
          </div>
          {candidates.length>0&&<button onClick={exportExcel} style={{
            padding:"6px 16px",borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer",
            background:"#1d4ed8",color:"#fff",border:"none",
          }}>Download Excel</button>}
          <button onClick={()=>loadCandidates(sel)} style={{
            padding:"6px 16px",borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer",
            background:"#fff",color:"#1d4ed8",border:"1px solid #1d4ed8",
          }}>{candLoading?"Loading...":"Refresh Data"}</button>
          <label style={{
            padding:"6px 16px",borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer",
            background:"#fff",color:"#059669",border:"1px solid #059669",display:"inline-block",
          }}>
            Upload CSV
            <input type="file" accept=".csv,.txt" style={{display:"none"}} onChange={e=>{if(e.target.files[0])handleUpload(e.target.files[0]);e.target.value="";}}/>
          </label>
        </div>
        {uploadStatus&&(
          <div style={{padding:"6px 12px",borderRadius:6,fontSize:11,marginBottom:8,
            background:uploadStatus.startsWith("Error")||uploadStatus.startsWith("Upload failed")?"#fef2f2":"#ecfdf5",
            color:uploadStatus.startsWith("Error")||uploadStatus.startsWith("Upload failed")?"#991b1b":"#065f46",
          }}>{uploadStatus}</div>
        )}

        {candidates.length>0?(
        <>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:8}}>
          <input type="text" placeholder="Search candidates..." value={candFilter} onChange={e=>setCandFilter(e.target.value)}
            style={{padding:"5px 10px",borderRadius:6,border:"1px solid #d1d5db",fontSize:12,flex:1,minWidth:140}}/>
          <select value={candLevelFilter} onChange={e=>setCandLevelFilter(e.target.value)} style={{
            padding:"5px 10px",borderRadius:6,border:"1px solid #d1d5db",fontSize:12,background:"#fff",
          }}>
            <option value="all">All Levels</option>
            <option value="federal">Federal</option>
            <option value="state">State</option>
            <option value="county">County</option>
            <option value="municipal">Municipal</option>
            <option value="school">School Board</option>
            <option value="special">Special District</option>
          </select>
        </div>

        <div style={{maxHeight:400,overflowY:"auto",border:"1px solid #e5e7eb",borderRadius:8}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead>
              <tr style={{background:"#f8fafc",position:"sticky",top:0}}>
                {[["name","Name"],["office","Office"],["party","Party"],
                  ...(!mobile?[["email","Email"],["phone","Phone"],["website","Website"]]:[])
                ].map(([key,label])=>(
                  <th key={key} onClick={()=>toggleSort(key)} style={{
                    padding:"6px 8px",textAlign:"left",fontWeight:700,borderBottom:"2px solid #e5e7eb",
                    color:sortCol===key?"#1d4ed8":"#374151",cursor:"pointer",userSelect:"none",
                  }}>{label} {sortCol===key?(sortAsc?"▲":"▼"):""}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredCandidates.slice(0,200).map((c,i)=>(
                <tr key={i} style={{borderBottom:"1px solid #f1f5f9"}}>
                  <td style={{padding:"5px 8px",fontWeight:600,color:"#0f172a"}}>{c.name}</td>
                  <td style={{padding:"5px 8px",color:"#4b5563",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.office}</td>
                  <td style={{padding:"5px 8px"}}>
                    <span style={{padding:"1px 8px",borderRadius:10,fontSize:10,fontWeight:700,
                      background:c.party==="DEM"?"#dbeafe":c.party==="REP"?"#fee2e2":"#f3f4f6",
                      color:c.party==="DEM"?"#1e40af":c.party==="REP"?"#991b1b":"#374151",
                    }}>{c.party||"—"}</span>
                  </td>
                  {!mobile&&<td style={{padding:"5px 8px",fontSize:11,color:c.email?"#059669":"#d1d5db"}}>{c.email||"—"}</td>}
                  {!mobile&&<td style={{padding:"5px 8px",fontSize:11,color:c.phone?"#059669":"#d1d5db"}}>{c.phone||"—"}</td>}
                  {!mobile&&<td style={{padding:"5px 8px",fontSize:11}}>{c.website?<a href={c.website.startsWith("http")?c.website:`https://${c.website}`} target="_blank" rel="noopener noreferrer" style={{color:"#1d4ed8"}}>{c.website.replace(/^https?:\/\//,"").slice(0,30)}</a>:<span style={{color:"#d1d5db"}}>—</span>}</td>}
                </tr>
              ))}
            </tbody>
          </table>
          {filteredCandidates.length>200&&(
            <div style={{padding:8,textAlign:"center",fontSize:11,color:"#6b7280",background:"#f8fafc"}}>
              Showing 200 of {filteredCandidates.length} — download Excel for full list
            </div>
          )}
        </div>
        <div style={{display:"flex",gap:12,marginTop:8,fontSize:11,color:"#6b7280"}}>
          <span>With email: {filteredCandidates.filter(c=>c.email).length}</span>
          <span>With phone: {filteredCandidates.filter(c=>c.phone).length}</span>
          <span>With website: {filteredCandidates.filter(c=>c.website).length}</span>
        </div>
        </>
        ):(
          <div style={{padding:16,textAlign:"center",color:"#6b7280",fontSize:12}}>
            {candLoading?"Loading candidate data...":"No candidate data yet. Click Refresh Data to pull from sources, or Upload CSV to add manually."}
          </div>
        )}
      </div>
    </div>
    )}
  </div>
  )}
  </div>
  );
}
export default App;
