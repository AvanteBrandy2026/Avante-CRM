import React, { useState, useEffect, useMemo, useRef } from 'react';
import { LayoutDashboard, Users, ClipboardList, Settings, TrendingUp, Phone, Mail, Search, Plus, X, ChevronRight, DollarSign, Award, Activity, Briefcase, Wine, ArrowUpRight, Save, RotateCcw, Target, BarChart3, Trash2, Download, FileSpreadsheet, UserPlus, Edit2 } from 'lucide-react';
import * as XLSX from 'xlsx';

const SEED_CLIENTS = [{"id":1,"venue":"The Lawns","channel":"Private Sales","firstName":"","lastName":"","location":"Camps Bay","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Low","lastContacted":"","status":"New","notes":"Need to make contact","totalSales":0},{"id":2,"venue":"50 on Gugs","channel":"Private Sales","firstName":"Sinbad","lastName":"","location":"Gugulethu","distributor":"Ultra","email":"","phone":"082 553 3599","leadSource":"Referral","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-18","status":"New","notes":"","totalSales":0},{"id":3,"venue":"Mitchell's","channel":"Private Sales","firstName":"Leroy","lastName":"","location":"CBD","distributor":"","email":"","phone":"660793944","leadSource":"Walk into Store","accountManager":"Alex","priority":"Medium","lastContacted":"2026-01-29","status":"New","notes":"Parking for now","totalSales":0},{"id":4,"venue":"Anthm","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":5,"venue":"Athletic Club & Social","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":6,"venue":"Hennies","channel":"Private Sales","firstName":"","lastName":"","location":"Stellenbosch","distributor":"","email":"","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"Medium","lastContacted":"2026-02-10","status":"New","notes":"Parking for now","totalSales":0},{"id":7,"venue":"Beyond","channel":"Private Sales","firstName":"Sebastian","lastName":"Stehr","location":"Constantia","distributor":"","email":"chef@beyondrestaurant.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-10","status":"New","notes":"Follow up in June","totalSales":0},{"id":8,"venue":"Rust en Vrede","channel":"Private Sales","firstName":"Fabio","lastName":"Daniel","location":"Stellenbosch","distributor":"","email":"fabio@rustenvrede.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-10","status":"Converted","notes":"Rejected","totalSales":0},{"id":9,"venue":"Blue Peter","channel":"Private Sales","firstName":"Marshal","lastName":"","location":"Blouberg","distributor":"","email":"","phone":"083 981 8770","leadSource":"Referral","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"Converted","notes":"","totalSales":0},{"id":10,"venue":"The Tasting Room at Creation","channel":"Private Sales","firstName":"Eleanor","lastName":"Niehaus","location":"Hermanus","distributor":"","email":"eleanor@creationwines.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-10","status":"New","notes":"Rejected","totalSales":0},{"id":11,"venue":"Dusk","channel":"Private Sales","firstName":"Callen","lastName":"Austin","location":"Stellenbosch","distributor":"","email":"","phone":"081 409 6544","leadSource":"Referral","accountManager":"Alex","priority":"Low","lastContacted":"2026-03-11","status":"Contacted","notes":"Rejected","totalSales":0},{"id":12,"venue":"Firemans Arms","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-20","status":"New","notes":"Promo finished. Not a focus.","totalSales":0},{"id":13,"venue":"Chefs Table","channel":"Private Sales","firstName":"Karl","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"2026-04-20","status":"Converted","notes":"Calling today","totalSales":0},{"id":14,"venue":"Bossa Group","channel":"Private Sales","firstName":"Chris","lastName":"","location":"Paarl","distributor":"","email":"","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-03-24","status":"Prospect","notes":"Parking for now","totalSales":0},{"id":15,"venue":"Chefs Warehouse Bree","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":16,"venue":"Aubergine","channel":"Private Sales","firstName":"Harald","lastName":"Bresselschmidt","location":"CBD","distributor":"","email":"chef@aubergine.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"Converted","notes":"Meeting on 24th","totalSales":0},{"id":17,"venue":"Chorus","channel":"Private Sales","firstName":"Mareli","lastName":"Basson","location":"Somerset West","distributor":"","email":"mareli@bertusbasson.com / info@bertusbasson.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Calling today","totalSales":0},{"id":18,"venue":"Club Heritage","channel":"Private Sales","firstName":"Pholoso","lastName":"","location":"","distributor":"","email":"","phone":"073 798 1918","leadSource":"Referral","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-17","status":"Converted","notes":"","totalSales":0},{"id":19,"venue":"Club Soho","channel":"Private Sales","firstName":"Krys","lastName":"","location":"","distributor":"","email":"djkrys86@gmail.com","phone":"083 533 9789","leadSource":"Referral","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"New","notes":"","totalSales":0},{"id":20,"venue":"Foxcroft","channel":"Private Sales","firstName":"Zoe","lastName":"Wepener","location":"","distributor":"","email":"beauadmin@chefswarehouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Emailed. Waiting for response.","totalSales":0},{"id":21,"venue":"Cyra","channel":"Private Sales","firstName":"Candice","lastName":"Philip","location":"Jo'Burg","distributor":"","email":"chefcandicephilip@gmail.com","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-03-10","status":"New","notes":"","totalSales":0},{"id":22,"venue":"Gigi","channel":"Private Sales","firstName":"Moses","lastName":"Moloi","location":"CBD","distributor":"","email":"moses@gigirestaurant.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"No response. Call.","totalSales":0},{"id":23,"venue":"Grand Cafe","channel":"Private Sales","firstName":"Rees","lastName":"Maxwell","location":"Sea Point","distributor":"","email":"","phone":"082 874 2299","leadSource":"Referral","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"Prospect","notes":"Struggles to get hold of Rees. Will keep persisting","totalSales":0},{"id":24,"venue":"Grub & Vne","channel":"Private Sales","firstName":"Matthew","lastName":"Manning","location":"","distributor":"","email":"info@mattmanningchef.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Emailed. Waiting for response.","totalSales":0},{"id":25,"venue":"H\u014dseki","channel":"Private Sales","firstName":"Virgil","lastName":"Kahn","location":"Stellenbosch","distributor":"","email":"lodge.chef@delaire.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"No response. Call.","totalSales":0},{"id":26,"venue":"La Colombe","channel":"Private Sales","firstName":"James","lastName":"Gaag","location":"Hout Bay","distributor":"","email":"james@lacolombe.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":27,"venue":"Post & Pepper","channel":"Private Sales","firstName":"Jess","lastName":"van Dyk","location":"Stellenbosch","distributor":"","email":"jess@postandpepper.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"Converted","notes":"Follow up in June","totalSales":0},{"id":28,"venue":"Serendipity","channel":"Private Sales","firstName":"Lizelle","lastName":"Stolze","location":"Wiilderness","distributor":"","email":"chef123@mweb.co.za/\nrestaurant@serendipitywilderness.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Call.","totalSales":0},{"id":29,"venue":"aha Hotels & Lodges","channel":"Private Sales","firstName":"Ruzandri","lastName":"","location":"","distributor":"","email":"ruzandri.stoltz@aha.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":30,"venue":"Hyatt Regency Cape Town","channel":"Private Sales","firstName":"Sasha","lastName":"Lewis","location":"","distributor":"","email":"sasha.lewis@hyatt.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":31,"venue":"Fac Cactus - Blouberg","channel":"Private Sales","firstName":"Lesley","lastName":"","location":"Blouberg","distributor":"","email":"simon@fatcactus.co.za","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"High","lastContacted":"2026-01-22","status":"New","notes":"","totalSales":0},{"id":32,"venue":"Silo (RP)","channel":"Private Sales","firstName":"Gerrit","lastName":"","location":"","distributor":"","email":"gerrit@trp.travel","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":33,"venue":"Steenberg GC","channel":"Private Sales","firstName":"Tatiana","lastName":"","location":"","distributor":"","email":"fbmanager@steenberggolfclub.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":34,"venue":"Twelve Apostles Hotel","channel":"Private Sales","firstName":"Mo","lastName":"Taliet","location":"","distributor":"","email":"groupproman12a@rchmail.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":35,"venue":"Westlake GC","channel":"Private Sales","firstName":"Tina","lastName":"","location":"","distributor":"","email":"events@westlakegolfclub.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":36,"venue":"De Zalze GC","channel":"Private Sales","firstName":"Ian","lastName":"","location":"","distributor":"","email":"FBManager@dezalzegolf.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":37,"venue":"Pearl Valley GC","channel":"Private Sales","firstName":"Stafford","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":38,"venue":"Arabella GC","channel":"Private Sales","firstName":"Sterlia","lastName":"Van Der Merwe","location":"","distributor":"","email":"sterlia.vandermerwe@southernsun.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":39,"venue":"Red Carnation Hotels","channel":"Private Sales","firstName":"Jameel","lastName":"","location":"","distributor":"","email":"foodbeverage@oysterbox.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":40,"venue":"Royal Johannesburg GC","channel":"Private Sales","firstName":"Erik","lastName":"","location":"","distributor":"","email":"fb@royaljhb.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":41,"venue":"Happy Folks","channel":"Private Sales","firstName":"Jacobus","lastName":"","location":"Stellenbosch / Blouberg","distributor":"","email":"blouberg@happyfolks.co.za","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"Low","lastContacted":"2026-01-20","status":"New","notes":"Need to send him a email and get Franchise approval to list Avante.","totalSales":0},{"id":42,"venue":"Madame Zingara Group","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":43,"venue":"More Family Collection","channel":"Private Sales","firstName":"Lara","lastName":"","location":"","distributor":"","email":"lara@more.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":44,"venue":"The Bay Hotel","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":45,"venue":"Premier Hotels","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":46,"venue":"Kaapstadt Brauhaus","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-21","status":"New","notes":"","totalSales":0},{"id":47,"venue":"Paarl Brandy Bar","channel":"Private Sales","firstName":"Allan","lastName":"","location":"Paarl","distributor":"","email":"","phone":"082 447 2538","leadSource":"Event","accountManager":"Alex","priority":"Low","lastContacted":"2026-04-02","status":"Qualified","notes":"Opening bar in Feb. Need to set up eeting this week","totalSales":0},{"id":48,"venue":"De Tafel","channel":"Private Sales","firstName":"Gregory","lastName":"Henderson","location":"Wynberg","distributor":"","email":"execchef@palmhouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"New","notes":"Waiting on response","totalSales":0},{"id":49,"venue":"Delaire Graff Restaurant","channel":"Private Sales","firstName":"Clinton","lastName":"Jacobs","location":"Stellenbosch","distributor":"","email":"dgr.headchef@delaire.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"New","notes":"No response. Call.","totalSales":0},{"id":50,"venue":"Fable","channel":"Private Sales","firstName":"Josh","lastName":"Sarembock","location":"CBD","distributor":"","email":"","phone":"060 313 4539","leadSource":"Event","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"Prospect","notes":"Keen to stock. Contracts end in June.","totalSales":0},{"id":51,"venue":"Fermier","channel":"Private Sales","firstName":"Adriaan","lastName":"Maree","location":"Pretoria","distributor":"","email":"bookings@fermierrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"New","notes":"No response. Call.","totalSales":0},{"id":52,"venue":"Salsify at the Roundhouse","channel":"Private Sales","firstName":"Ryan","lastName":"Cole","location":"Camps Bay","distributor":"","email":"ryan@salsify.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Figuring out distribution","totalSales":0},{"id":53,"venue":"Les Cr\u00e9atifs","channel":"Private Sales","firstName":"Wandile","lastName":"Mabaso","location":"Jo'Burg","distributor":"","email":"chefwmabaso@gmail.com","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-10","status":"Converted","notes":"Meeting end of May","totalSales":0},{"id":54,"venue":"Terrarium","channel":"Private Sales","firstName":"Anlou","lastName":"Erasmus","location":"Waterfront","distributor":"","email":"qvchef@queenvictoriahotel.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":55,"venue":"The Chefs' Table","channel":"Private Sales","firstName":"Mathew","lastName":"Armbruster","location":"","distributor":"","email":"mathew@thechefstable.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":56,"venue":"The Jordan Restaurant with Marthinus Ferreira","channel":"Private Sales","firstName":"Leigh","lastName":"Simmidari","location":"Stellenbosch","distributor":"","email":"restaurant@jordanwines.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Calling today","totalSales":0},{"id":57,"venue":"Upper Union","channel":"Private Sales","firstName":"Amori","lastName":"Burger","location":"CBD","distributor":"","email":"amori@upperunion.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"Converted","notes":"Followed up. Waiting.","totalSales":0},{"id":58,"venue":"Chefs Warehouse Maison","channel":"Private Sales","firstName":"David Schneider","lastName":"Xavier Francis (head chef)","location":"Franschoek","distributor":"","email":"david@chefswarehouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-06","status":"New","notes":"No response. Call.","totalSales":0},{"id":59,"venue":"Medusa","channel":"Private Sales","firstName":"Vuyani","lastName":"","location":"Milnerton","distributor":"","email":"medusaloungerooftop@gmail.com","phone":"072 258 2463","leadSource":"Cold Call","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-16","status":"Converted","notes":"Done Deals","totalSales":0},{"id":60,"venue":"MIlnerton Golf Club","channel":"Private Sales","firstName":"Alfie","lastName":"Schneeburger","location":"Milnerton","distributor":"","email":"catering@milgolf.co.za","phone":"082 576 7887","leadSource":"Event","accountManager":"Loydz","priority":"Low","lastContacted":"2026-01-20","status":"New","notes":"","totalSales":0},{"id":61,"venue":"Misthom Lounge","channel":"Private Sales","firstName":"Lee","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"Low","lastContacted":"2026-02-20","status":"New","notes":"","totalSales":0},{"id":62,"venue":"Kove Collection","channel":"Private Sales","firstName":"Benito","lastName":"","location":"Camps Bay","distributor":"Karino","email":"","phone":"063 212 7038","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-07","status":"Prospect","notes":"Benito likes Avante. We have a relationship. Waiting for an order and which restaurants.","totalSales":0},{"id":63,"venue":"Mozambik","channel":"Private Sales","firstName":"Uhuru","lastName":"","location":"","distributor":"","email":"","phone":"084 057 0572","leadSource":"Walk into Store","accountManager":"Loydz","priority":"High","lastContacted":"2026-01-21","status":"New","notes":"","totalSales":0},{"id":64,"venue":"Noah","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":65,"venue":"Eike","channel":"Private Sales","firstName":"Michael","lastName":"Fuller","location":"Stellenbosch","distributor":"","email":"kitchen@eikerestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"Callling today","totalSales":0},{"id":66,"venue":"FABER","channel":"Private Sales","firstName":"Dale","lastName":"Stevens","location":"Paarl","distributor":"","email":"faber@avondalewine.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"No response. Call.","totalSales":0},{"id":67,"venue":"TTK Fledgelings","channel":"Private Sales","firstName":"Nathan","lastName":"Clarke","location":"CBD","distributor":"","email":"steven@lukedaleroberts.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":68,"venue":"Life and Brand Portfolo","channel":"Private Sales","firstName":"Jason","lastName":"","location":"Cape Town","distributor":"","email":"jason@lifeandbrand.co.za","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-08","status":"New","notes":"Waiting for response","totalSales":0},{"id":69,"venue":"Spice Route","channel":"Private Sales","firstName":"Gigi","lastName":"Bisogno","location":"Paarl","distributor":"","email":"","phone":"083 441 9870","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-08","status":"Contacted","notes":"","totalSales":0},{"id":70,"venue":"Qunu","channel":"Private Sales","firstName":"Matthew","lastName":"Foxon","location":"Jo'Burg","distributor":"","email":"matthewf@saxon.co.za/scottd@saxon.co.za","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-10","status":"New","notes":"Given to Loydz","totalSales":0},{"id":71,"venue":"Room13, Savage & Concept","channel":"Private Sales","firstName":"Prince","lastName":"","location":"","distributor":"","email":"","phone":"061 982 6349","leadSource":"Referral","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"New","notes":"","totalSales":0},{"id":72,"venue":"La Petite Colombe","channel":"Private Sales","firstName":"Peter","lastName":"Duncan","location":"Franschoek","distributor":"","email":"peter@lapetitecolombe.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-08","status":"Lost","notes":"Revisit in Q2","totalSales":0},{"id":73,"venue":"Le coin Fran\u00e7ais","channel":"Private Sales","firstName":"Darren","lastName":"Badenhorst","location":"Franschoek","distributor":"","email":"darren@lecoinfrancais.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-08","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":74,"venue":"PIER","channel":"Private Sales","firstName":"John","lastName":"Norris-Rogers","location":"Waterfront","distributor":"","email":"john@pier.restaurant","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-08","status":"New","notes":"Callling today","totalSales":0},{"id":75,"venue":"Seat - Blouberg","channel":"Private Sales","firstName":"Anton","lastName":"","location":"Blouberg","distributor":"","email":"info@seatrestaurant.com","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-01-22","status":"New","notes":"","totalSales":0},{"id":76,"venue":"LifeGrande Group","channel":"Private Sales","firstName":"Jeanel / John","lastName":"","location":"","distributor":"","email":"jeanel@lifegrandcafe.com / john@lifegrandcafe.com","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-09","status":"Prospect","notes":"Waiting for response","totalSales":0},{"id":77,"venue":"Stellenbosch GC","channel":"Private Sales","firstName":"Gerhard","lastName":"","location":"Stellenbosch","distributor":"Karino","email":"chef@stbgolf.com","phone":"072 198 7912","leadSource":"Event","accountManager":"Alex","priority":"High","lastContacted":"2026-04-09","status":"Prospect","notes":"Waitng for order coonfirmation","totalSales":0},{"id":78,"venue":"Galjoen","channel":"Private Sales","firstName":"Isca","lastName":"Stoltz","location":"CBD","distributor":"","email":"eat@galjoencpt.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-09","status":"New","notes":"No response. Call.","totalSales":0},{"id":79,"venue":"UCT Bar","channel":"Private Sales","firstName":"Rob","lastName":"Munroe","location":"UCT","distributor":"","email":"","phone":"079 569 4688","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-13","status":"Converted","notes":"Confirmed","totalSales":0},{"id":80,"venue":"SoHo","channel":"Private Sales","firstName":"DJ Krys","lastName":"","location":"Long Street","distributor":"","email":"083 533 9789","phone":"083 533 9789","leadSource":"Referral","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-21","status":"Converted","notes":"","totalSales":0},{"id":81,"venue":"Belly of the Beast","channel":"Private Sales","firstName":"Anouchka","lastName":"Horn","location":"CBD","distributor":"","email":"eat@bellyofthebeast.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-13","status":"New","notes":"No response. Call.","totalSales":0},{"id":82,"venue":"Cavalli","channel":"Private Sales","firstName":"Lucas","lastName":"Carstens","location":"Stellenbosch","distributor":"","email":"chef@cavalliestate.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-13","status":"New","notes":"Calling today","totalSales":0},{"id":83,"venue":"Embarc","channel":"Private Sales","firstName":"Darren","lastName":"O'Donovan","location":"Jo'Burg","distributor":"","email":"info@embarcrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-13","status":"New","notes":"No response. Call.","totalSales":0},{"id":84,"venue":"Thirsty Scarecrow","channel":"Private Sales","firstName":"Ivan","lastName":"Botha","location":"Stellenbosch","distributor":"Direct","email":"kzetler@gmail.com","phone":"084 263 4243","leadSource":"Event","accountManager":"Alex","priority":"Low","lastContacted":"2026-04-14","status":"Converted","notes":"Need to drop barrels. Waiting for order amount.","totalSales":0},{"id":85,"venue":"&Beyond","channel":"Private Sales","firstName":"Nicole Robinson","lastName":"","location":"","distributor":"","email":"nicole.robinson@andbeyond.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-14","status":"New","notes":"","totalSales":0},{"id":86,"venue":"Talking to Strangers","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":87,"venue":"Wolfgat","channel":"Private Sales","firstName":"Kobus","lastName":"van der Merwe","location":"","distributor":"","email":"jjvandermerwe@gmail.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-15","status":"New","notes":"Calling today","totalSales":0},{"id":88,"venue":"Thornybush","channel":"Private Sales","firstName":"Zanele","lastName":"","location":"","distributor":"","email":"","phone":"157931976","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-15","status":"New","notes":"","totalSales":0},{"id":89,"venue":"The Waterside","channel":"Private Sales","firstName":"Roxy","lastName":"Mudie","location":"Waterfront","distributor":"","email":"roxy@thewaterside.restaurant","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-16","status":"New","notes":"Calling today","totalSales":0},{"id":90,"venue":"Clovelly Golf Club","channel":"Private Sales","firstName":"Paul","lastName":"","location":"","distributor":"","email":"gm@clovelly.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-16","status":"New","notes":"","totalSales":0},{"id":91,"venue":"The LivingRoom","channel":"Private Sales","firstName":"Johannes","lastName":"Richter","location":"Durban","distributor":"","email":"info@summerhillkzn.com","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-01-22","status":"New","notes":"Given to Loydz","totalSales":0},{"id":92,"venue":"The Pot Luck Club, Johannesburg","channel":"Private Sales","firstName":"Ebie du Toit/ Vicky Peech","lastName":"Ebie du Toit/ Vicky Peech","location":"Jo'Burg","distributor":"","email":"ebi@plcj.co.za traveller@thepeech.co.za","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-01-22","status":"New","notes":"","totalSales":0},{"id":93,"venue":"Marriott / Protea Hotels","channel":"Private Sales","firstName":"Unathi","lastName":"Dyonase","location":"","distributor":"","email":"Unathi.dyonase@marriot.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-17","status":"New","notes":"","totalSales":0},{"id":94,"venue":"Chefs Warehouse Beau Constantia","channel":"Private Sales","firstName":"Ivor","lastName":"Jones","location":"Constantia","distributor":"","email":"ivor@chefswarehouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-20","status":"New","notes":"Followed up in March","totalSales":0},{"id":95,"venue":"COY","channel":"Private Sales","firstName":"Teenola","lastName":"Govender","location":"CBD","distributor":"","email":"teenola@coyrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-20","status":"New","notes":"Callling today","totalSales":0},{"id":96,"venue":"The Red Room","channel":"Private Sales","firstName":"Caroline","lastName":"Lamb","location":"CBD","distributor":"","email":"caroline@cwredroom.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-21","status":"Converted","notes":"Meeting next week","totalSales":0},{"id":97,"venue":"Tsogo Sun Mossel Bay","channel":"Private Sales","firstName":"George","lastName":"Webber","location":"Mossel Bay","distributor":"","email":"","phone":"838227550","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-21","status":"New","notes":"","totalSales":0},{"id":98,"venue":"\u00eblgr.","channel":"Private Sales","firstName":"Jesper","lastName":"Nillson","location":"CBD","distributor":"","email":"jesper@elgr.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-22","status":"New","notes":"Callling today","totalSales":0},{"id":99,"venue":"FYN","channel":"Private Sales","firstName":"Peter","lastName":"Templehoff","location":"CBD","distributor":"","email":"tempelhoff@gmail.com/ Kerry@fynrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-22","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":100,"venue":"Quay 4","channel":"Private Sales","firstName":"Leroy","lastName":"","location":"Waterfront","distributor":"","email":"","phone":"660793944","leadSource":"Walk into Store","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-22","status":"New","notes":"Waiting response for a meeting with owner","totalSales":0},{"id":101,"venue":"Southern Sun","channel":"Private Sales","firstName":"David","lastName":"Mokwebo","location":"","distributor":"","email":"david.mokwebo@southernsun.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"","totalSales":0},{"id":102,"venue":"Unclaimed","channel":"Private Sales","firstName":"Aidan","lastName":"","location":"Kloof Street","distributor":"","email":"","phone":"079 380 9210","leadSource":"Networking","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"Converted","notes":"","totalSales":0},{"id":103,"venue":"BMW Dealership","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"","totalSales":0},{"id":104,"venue":"Bidvest Lounge","channel":"Private Sales","firstName":"Suzanne","lastName":"VD","location":"","distributor":"","email":"SuzanneVD@bidvestcatering.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"Sent to procurement department","totalSales":0},{"id":105,"venue":"Molenvlliet","channel":"Private Sales","firstName":"Juliene","lastName":"","location":"","distributor":"","email":"stay@molenvliet.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":106,"venue":"Babylonstoren","channel":"Private Sales","firstName":"Elana","lastName":"","location":"","distributor":"","email":"functions@babylonstoren.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"High","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":107,"venue":"Boschendal","channel":"Private Sales","firstName":"Shanel","lastName":"","location":"","distributor":"","email":"functions@babylonstoren.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"High","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":108,"venue":"Da'Aria","channel":"Private Sales","firstName":"Danielle","lastName":"","location":"","distributor":"","email":"events@daria.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"High","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":109,"venue":"Slow Lounge","channel":"Private Sales","firstName":"Janine","lastName":"","location":"","distributor":"","email":"janine@purefoods.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"Waiting for response","totalSales":0},{"id":110,"venue":"Youngblood","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":111,"venue":"OK Liquors Parow","channel":"Trade Retail","firstName":"Denver","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"844552084","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"New","notes":"Will talk to owner and get back","totalSales":0},{"id":112,"venue":"OK Liquors Tygerdal","channel":"Trade Retail","firstName":"Ray","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"211090099","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":113,"venue":"Tops Bellville","channel":"Trade Retail","firstName":"Yushrah","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"638150626","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":114,"venue":"Tops Boston","channel":"Trade Retail","firstName":"Valencia","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"653915424","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":115,"venue":"Tops Glenwood","channel":"Trade Retail","firstName":"Obert","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"787540057","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":116,"venue":"Tops Parow Valley","channel":"Trade Retail","firstName":"Fiona","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"746320000","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":117,"venue":"Observatory Bottle Store","channel":"Trade Retail","firstName":"Dani","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"719445117","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-16","status":"New","notes":"Not interested yet","totalSales":0},{"id":118,"venue":"Blue Bottle Observatory","channel":"Trade Retail","firstName":"Ling","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-16","status":"New","notes":"Owner will get back to me","totalSales":0},{"id":119,"venue":"Blue Bottle Panorama (Day 3)","channel":"Trade Retail","firstName":"Carl","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"842054604","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":120,"venue":"Market Liquors Parklands Sandowne","channel":"Trade Retail","firstName":"Adriana","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"822325842","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":121,"venue":"Tops Belhar","channel":"Trade Retail","firstName":"Jonothan","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"673236926","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-13","status":"New","notes":"Sent owner a mail, no feedback","totalSales":0},{"id":122,"venue":"Blue Bottle Wild Liquors","channel":"Trade Retail","firstName":"Evan","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"745206017","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":123,"venue":"Liquor City Paarl","channel":"Trade Retail","firstName":"Zheng","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"766767666","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":124,"venue":"Market Liquors Malmesbury","channel":"Trade Retail","firstName":"Janicka","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"840444024","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":125,"venue":"Market Liquors Paarl (Day 1)","channel":"Trade Retail","firstName":"Joy","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"672330172","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":126,"venue":"OK Liquors Rembrandt Mall","channel":"Trade Retail","firstName":"Burton","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"New","notes":"Under new ownership","totalSales":0},{"id":127,"venue":"Tops Zomerlust (Day 1)","channel":"Trade Retail","firstName":"Dimitri","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"769262883","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":128,"venue":"Blue Bottle Liquors Paarl Mall","channel":"Trade Retail","firstName":"Louise","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-02","status":"New","notes":"Not interested yet","totalSales":0},{"id":129,"venue":"Blue Bottle Liquor for Africa","channel":"Trade Retail","firstName":"Yolande","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"836072111","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":130,"venue":"Market Liquors Brackenfell","channel":"Trade Retail","firstName":"Elizabeth","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"612727071","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":131,"venue":"OK Liquors Goedemoed","channel":"Trade Retail","firstName":"Dricus","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"719122379","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":132,"venue":"Tops Brackenfell","channel":"Trade Retail","firstName":"Johan","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"652973469","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":133,"venue":"Tops Kraaifontein (Day 8)","channel":"Trade Retail","firstName":"Adriaan","lastName":"","location":"Cape Town","distributor":"Karino","email":"kraaifonteinsuperspar@retail.co.za","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":134,"venue":"Tops Protea Hoogte (Day 8)","channel":"Trade Retail","firstName":"LJ","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"712375157","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":135,"venue":"Tops Helshoogte","channel":"Trade Retail","firstName":"Patuma","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"763409494","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":136,"venue":"Carolines Tokai","channel":"Trade Retail","firstName":"Caroline","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"New","notes":"Not interested","totalSales":0},{"id":137,"venue":"Market Liquors Hermanus","channel":"Trade Retail","firstName":"JJ","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"798977999","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":138,"venue":"OK Liquors Napier","channel":"Trade Retail","firstName":"Ben","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"822576721","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"New","notes":"No feedback whatsoever","totalSales":0},{"id":139,"venue":"Tops Kleinmond (Day 9)","channel":"Trade Retail","firstName":"Marjana","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"798331149","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":140,"venue":"Tops Villiersdorp (Day 9)","channel":"Trade Retail","firstName":"Elaine","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"662807287","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":141,"venue":"Tops Welgevonden (Day 7)","channel":"Trade Retail","firstName":"Wayne","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"621274008","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":142,"venue":"Market Liquors Tokai (Day 10)","channel":"Trade Retail","firstName":"Kim","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"730617906","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":143,"venue":"OK Liquors Spineroad","channel":"Trade Retail","firstName":"Eleanor","lastName":"","location":"Cape Town","distributor":"Karino","email":"spineroad@ok.co.za","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"New","notes":"Under new ownership","totalSales":0},{"id":144,"venue":"Tops Fishhoek","channel":"Trade Retail","firstName":"Gilbert","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"768108940","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"Listed VS & VSOP","totalSales":0},{"id":145,"venue":"Tops Glencairn","channel":"Trade Retail","firstName":"Theresa","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Cold Call","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":146,"venue":"Tops Houtbay","channel":"Trade Retail","firstName":"Farai","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"710779622","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":147,"venue":"Tops Muizenberg","channel":"Trade Retail","firstName":"Denise","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"820525529","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":148,"venue":"Tops Oakhurst","channel":"Trade Retail","firstName":"Chicco","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"609232264","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":149,"venue":"Tops Weltevreden (Day 10)","channel":"Trade Retail","firstName":"Paul","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"610052357","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":150,"venue":"Blue Bottle Elands Bay","channel":"Trade Retail","firstName":"Patrick","lastName":"","location":"West-Coast","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-23","status":"New","notes":"Will give feedback in April","totalSales":0},{"id":151,"venue":"Tops Edgemead","channel":"Trade Retail","firstName":"Nolene","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"724245680","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-05","status":"New","notes":"Not interested yet","totalSales":0},{"id":152,"venue":"Tops Bredasdorp","channel":"Trade Retail","firstName":"Anna","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"728264575","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-03","status":"Converted","notes":"","totalSales":0},{"id":153,"venue":"Tops Aurora (Day 6)","channel":"Trade Retail","firstName":"Alcino","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"765381927","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":154,"venue":"Tops Durbanville","channel":"Trade Retail","firstName":"Devon","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"815151023","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":155,"venue":"Tops Haasendal","channel":"Trade Retail","firstName":"Tyrese","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"677984010","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":156,"venue":"Tops Old Oak","channel":"Trade Retail","firstName":"Hennie","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"219192711","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":157,"venue":"Tops Palm Grove","channel":"Trade Retail","firstName":"Charl","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"219762148","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":158,"venue":"Tops Phesantekraal","channel":"Trade Retail","firstName":"Jestin","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"605484307","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":159,"venue":"Tops Kuilsriver (Day 5)","channel":"Trade Retail","firstName":"Kevin","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"785021144","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":160,"venue":"Tops Sunrise Circle (Day 5)","channel":"Trade Retail","firstName":"Barry","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"738878330","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":161,"venue":"Tops Mowbray","channel":"Trade Retail","firstName":"Michelle","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"765520284","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":162,"venue":"Tops Observatory","channel":"Trade Retail","firstName":"Lincon","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"658430389","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":163,"venue":"Tops Wex1","channel":"Trade Retail","firstName":"Ernst","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"738837072","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":164,"venue":"Tops Woodstock Quarter","channel":"Trade Retail","firstName":"Ernst","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"738837072","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":165,"venue":"Whisky Shop","channel":"Trade Retail","firstName":"Hector","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"833771113","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":166,"venue":"Liquor City Long Street (Day 4)","channel":"Trade Retail","firstName":"Arthur","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"680530467","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":167,"venue":"Woodstock Liquors","channel":"Trade Retail","firstName":"Inacio","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"832689430","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":168,"venue":"Tops Adderly","channel":"Trade Retail","firstName":"Noni","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":169,"venue":"Tops Cape Station","channel":"Trade Retail","firstName":"Noni","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":170,"venue":"Tops Century Village","channel":"Trade Retail","firstName":"Anthony","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"825215452","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":171,"venue":"Blouberg Liquors","channel":"Trade Retail","firstName":"Inacio","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"832689430","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":172,"venue":"NGF Sunset Beach","channel":"Trade Retail","firstName":"Gary","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"726145099","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":173,"venue":"OK Liquors Duynefontein","channel":"Trade Retail","firstName":"Frean","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"722437068","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":174,"venue":"Tops Blouberg","channel":"Trade Retail","firstName":"Canais","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"635074577","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":175,"venue":"Tops Melkbos","channel":"Trade Retail","firstName":"Abigail","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"715768900","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":176,"venue":"Tops Milnerton","channel":"Trade Retail","firstName":"Cheslyn","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"692649715","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":177,"venue":"Tops Morningfield (Day 3)","channel":"Trade Retail","firstName":"Sinetemba","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"833584184","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":178,"venue":"Tops Parklands","channel":"Trade Retail","firstName":"Morne","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"743749296","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":179,"venue":"Tops Riberios","channel":"Trade Retail","firstName":"Louise","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"673611248","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":180,"venue":"Tops Royal Ascot","channel":"Trade Retail","firstName":"Tumeka","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"632692784","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":181,"venue":"Tops Sunningdale","channel":"Trade Retail","firstName":"Tinotenda","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"658334199","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":182,"venue":"Tops Tableview","channel":"Trade Retail","firstName":"Roger","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"718921187","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":183,"venue":"Tops Malmesbury","channel":"Trade Retail","firstName":"Hendika","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"785751687","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":184,"venue":"Tops Vineyard","channel":"Trade Retail","firstName":"Christiaan","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"720407583","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":185,"venue":"Tops Cape Quarter (Day 2)","channel":"Trade Retail","firstName":"Chantal","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"836613653","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":186,"venue":"Barkeeper","channel":"Trade Retail","firstName":"Carla","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"823391079","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":187,"venue":"Constantia Wine & Craft","channel":"Trade Retail","firstName":"Jeremy","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"786424044","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":188,"venue":"NGF Gardens","channel":"Trade Retail","firstName":"Vuyo","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"7433425032","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":189,"venue":"NGF Three Anchor Bay","channel":"Trade Retail","firstName":"Joline","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"786584312","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":190,"venue":"Tops Bergvliet","channel":"Trade Retail","firstName":"Vuyi","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"766611750","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"Not interested yet","totalSales":0},{"id":191,"venue":"Tops Kloof","channel":"Trade Retail","firstName":"Makhumalo","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"631053550","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":192,"venue":"Tops Rosmead (Day 2)","channel":"Trade Retail","firstName":"Rochelle","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"728603028","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":193,"venue":"Tops Seapoint","channel":"Trade Retail","firstName":"Stuart","lastName":"","location":"CBD","distributor":"Karino","email":"seapoint1@retail.spar.co.za","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":194,"venue":"Tops Vredehoek","channel":"Trade Retail","firstName":"Lisa","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"682316301","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":195,"venue":"Tops Bottelary (Day 7)","channel":"Trade Retail","firstName":"Jestin","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"605484307","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-27","status":"Converted","notes":"","totalSales":0},{"id":196,"venue":"OK Liquors Urban Sonstraal","channel":"Trade Retail","firstName":"Bri-Niel","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"733719691","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":197,"venue":"Tops Cape Gate","channel":"Trade Retail","firstName":"Frank","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"638538885","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":198,"venue":"Liquor City Somerset Mall","channel":"Trade Retail","firstName":"Charlene","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"715243228","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":199,"venue":"OK Liquors Strand","channel":"Trade Retail","firstName":"Meaghan","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"733695375","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":200,"venue":"Tops Cinnamon","channel":"Trade Retail","firstName":"Tafadswa","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"614817661","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":201,"venue":"Tops De Jonker","channel":"Trade Retail","firstName":"Charlton","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"641629569","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":202,"venue":"Tops Die Boord","channel":"Trade Retail","firstName":"Gershwin","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"613656818","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":203,"venue":"Tops Helderbosch","channel":"Trade Retail","firstName":"Chantal","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"674027744","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":204,"venue":"Tops Helderveu","channel":"Trade Retail","firstName":"Heinrich","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"619525640","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":205,"venue":"Tops Lions Square","channel":"Trade Retail","firstName":"Barbara","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"790973527","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":206,"venue":"Tops Mountainview","channel":"Trade Retail","firstName":"Melisia","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"740583958","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":207,"venue":"Tops Paradyskloof","channel":"Trade Retail","firstName":"Kenny","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"739506767","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":208,"venue":"Tops Paul Roos","channel":"Trade Retail","firstName":"Nathely","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"780743365","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":209,"venue":"Tops Strand","channel":"Trade Retail","firstName":"Lycken","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"620494468","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"New","notes":"Not interesteed yet","totalSales":0},{"id":210,"venue":"Tops Twin Palms","channel":"Trade Retail","firstName":"Senobia","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"843573643","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":211,"venue":"Liquor City Claremont","channel":"Trade Retail","firstName":"Derek","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"216741478","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":212,"venue":"OK Liquors Ridgeworth","channel":"Trade Retail","firstName":"Quinton","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"837799296","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":213,"venue":"Picardi Rebel Claremont (Day 4)","channel":"Trade Retail","firstName":"Tristian","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"662086625","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":214,"venue":"Tops Eastcliff","channel":"Trade Retail","firstName":"Nadia","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"790927628","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":215,"venue":"Tops Eersteriver (Day 6)","channel":"Trade Retail","firstName":"Sinalo","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"735920264","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"New","notes":"Not interested at all","totalSales":0},{"id":216,"venue":"OK Liquors Strandfontein","channel":"Trade Retail","firstName":"Jenique","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"726130092","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":217,"venue":"Tops Alphen","channel":"Trade Retail","firstName":"Riyaaz","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"742624147","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":218,"venue":"Ultra Nyanga","channel":"Trade Retail","firstName":"Sisa","lastName":"Liwani","location":"Nyanga","distributor":"","email":"nyanga@ultraliquors.co.za","phone":"082 398 0404","leadSource":"Online","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-18","status":"New","notes":"","totalSales":0},{"id":219,"venue":"Tops Sonstraal","channel":"Trade Retail","firstName":"Selvinia","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"634486222","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":220,"venue":"Tops Vredekloof","channel":"Trade Retail","firstName":"Riaan","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"605905596","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":221,"venue":"Tops Welgelegeen","channel":"Trade Retail","firstName":"Yianakis","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"720615896","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":222,"venue":"Tops Zevenwacht","channel":"Trade Retail","firstName":"Wesley","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"824777793","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":223,"venue":"Dal Italia Restaurant","channel":"Trade Retail","firstName":"Fabio","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"823372339","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-06","status":"Converted","notes":"","totalSales":0},{"id":224,"venue":"City Lodge Hotels","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":225,"venue":"Anew Hotels","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":226,"venue":"Bon Hotels","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":227,"venue":"Steyn City","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":228,"venue":"Taj Hotel","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":229,"venue":"Glendower","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":230,"venue":"Randpark Golf Club","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":231,"venue":"Houghton Golf Club","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":232,"venue":"Londolozi","channel":"Private Sales","firstName":"Duncan","lastName":"","location":"","distributor":"","email":"duncan@londolozi.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"No response, call today","totalSales":0},{"id":233,"venue":"MalaMala","channel":"Private Sales","firstName":"Rufie or Ross","lastName":"","location":"","distributor":"","email":"","phone":"781008810","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"No response, call today","totalSales":0},{"id":234,"venue":"Shamwari Group","channel":"Private Sales","firstName":"Janine","lastName":"","location":"","distributor":"","email":"procurement@shamwari.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"","totalSales":0},{"id":235,"venue":"Singita (Group)","channel":"Private Sales","firstName":"Schwess","lastName":"","location":"","distributor":"","email":"capepremierwine@singita.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"","totalSales":0},{"id":236,"venue":"Foodbarn Group","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":237,"venue":"Hussar Grill Group","channel":"Private Sales","firstName":"Bradley","lastName":"","location":"","distributor":"","email":"bradleyh@spurcorp.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":238,"venue":"Newmark Hotels","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":239,"venue":"Burkenhead House (RP)","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":240,"venue":"Ellerman House","channel":"Private Sales","firstName":"Manuel","lastName":"","location":"","distributor":"","email":"manuel@ellerman.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":241,"venue":"Fairmont / Accor","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":242,"venue":"Hilton SA","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":243,"venue":"Kapama Private Game Reserve","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":244,"venue":"La Residence (RP)","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":245,"venue":"One&Only Cape Town","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":246,"venue":"Royal Malewane (RP)","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":247,"venue":"Sun International","channel":"Private Sales","firstName":"Mpho","lastName":"Moshotle","location":"","distributor":"","email":"Mpho.moshotle@suninternational.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":248,"venue":"Ulusaba (Virgin)","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":249,"venue":"Energy Masterbuilders Paarl","channel":"Trade Retail","firstName":"Johan","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":250,"venue":"Hoekstra Farming","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":251,"venue":"Nexus Agriculture","channel":"Trade Retail","firstName":"Ester","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":252,"venue":"PSG Paarl","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":253,"venue":"PSG Stellenbosch","channel":"Trade Retail","firstName":"Joaan","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":254,"venue":"BMW Paarlberg","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":255,"venue":"BMW Stellenbosch","channel":"Trade Retail","firstName":"Christopher","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":256,"venue":"GWM Paarl","channel":"Trade Retail","firstName":"Brandon","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":257,"venue":"Exsa","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":258,"venue":"Paarl Boys High Waterpolo","channel":"Trade Retail","firstName":"Ross","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":259,"venue":"Stellenberg High School","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":260,"venue":"New Orleans High School","channel":"Trade Retail","firstName":"Carlene","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":261,"venue":"Nuweland Paarl","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":262,"venue":"Giflo Group","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":263,"venue":"Samsung","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":264,"venue":"Icon Fruit","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":265,"venue":"Isipani Paarl","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":266,"venue":"Sky vines","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":267,"venue":"Akura Paarl","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":268,"venue":"Rawson H/Q","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":269,"venue":"Workshop 17","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":270,"venue":"Cargokwik","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":271,"venue":"Zuus Paarl","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":272,"venue":"Selavie (Val de Vie)","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":273,"venue":"Blacksmith Kitchen","channel":"Trade Retail","firstName":"Bianca","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":274,"venue":"Grande Roche Hotel","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":275,"venue":"Noop","channel":"Trade Retail","firstName":"Zian","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":276,"venue":"Laborie Estate","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":277,"venue":"Veld and Vine","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":278,"venue":"Val de Vie","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0}];
const LOGO_URL = "https://raw.githubusercontent.com/AvanteBrandy2026/Avante-CRM/main/public/logo.jpg"

const SALES_REPS = ['Alex', 'Lehmarc', 'Loydz'];

// Sales rep email addresses — used in the signature of order confirmation emails
// so the client knows who the message is from and where to reply.
const REP_EMAILS = {
  Alex: 'alex@breakfreebeverages.com',
  Lehmarc: 'lehmarc@breakfreebeverages.com',
  Loydz: 'lonwabo@breakfreebeverages.com',
};

// Default recipients for every order confirmation email — copied into the
// SEND TO field automatically. Reps can edit/add/remove before sending.
const DEFAULT_ORDER_RECIPIENTS = [
  'lauren@redbev.co.za',
  'melanie@redbev.co.za',
];
const CHANNELS = ['Private Sales', 'Trade Retail'];
const STATUSES = ['New', 'Contacted', 'Qualified', 'Prospect', 'Converted', 'Lost'];
const LEAD_SOURCES = ['Cold Call', 'Referral', 'Walk into Store', 'Email', 'Event', 'Networking', 'Call-cycle', 'Website', 'Online'];
const PRIORITIES = ['Low', 'Medium', 'High'];

const DEFAULT_TARGETS = {
  Alex: { revenue: 150000, visits: 60, privateSales: 40, tradeRetail: 20 },
  Lehmarc: { revenue: 150000, visits: 60, privateSales: 20, tradeRetail: 40 },
  Loydz: { revenue: 100000, visits: 45, privateSales: 30, tradeRetail: 15 },
};

// Avante SKU catalogue (Trade Ex VAT prices, ZAR per unit)
const SKU_CATALOGUE = [
  { id: 'vsop_750_current', name: 'VSOP 750ml (Current)', price: 420.00 },
  { id: 'vsop_750_new', name: 'VSOP 750ml (New)', price: 579.00 },
  { id: 'vsop_200', name: 'VSOP 200ml', price: 150.00 },
  { id: 'xv_750', name: 'XV 750ml', price: 2086.00 },
  { id: 'vs_500', name: 'VS 500ml', price: 217.00 },
  { id: 'vs_750', name: 'VS 750ml', price: 296.50 },
  { id: 'xv_200', name: 'XV 200ml', price: 745.00 },
  { id: 'xo_750', name: 'XO 750ml', price: 1043.00 },
  { id: 'vs_200', name: 'VS 200ml', price: 113.00 },
  { id: 'gift_4x50', name: '4 x 50ml Gift Box', price: 520.00 },
  { id: 'gift_3x200', name: '3 x 200ml Gift Box', price: 1050.00 },
  { id: 'custom_vs', name: 'Custom VS', price: 694.78 },
  { id: 'custom_xv', name: 'Custom XV', price: 694.78 },
];

const ZAR = (n) => 'R ' + Math.round(n || 0).toLocaleString('en-ZA');
const todayISO = () => new Date().toISOString().slice(0, 10);
const monthISO = () => new Date().toISOString().slice(0, 7);

// =================== Avante Logo ===================
// Logo is landscape 595x417 — use height + auto width
const AvanteLogo = ({ className = '', height = 52 }) => (
  <img
    src={LOGO_URL}
    alt="Avante Cape Brandy"
    className={className}
    style={{ height: height, width: 'auto', objectFit: 'contain', display: 'block', maxWidth: height * 1.8 }}
    onError={(e) => { e.target.style.display = 'none'; }}
  />
);

const AvanteWordmark = ({ dark = false, className = '' }) => (
  <svg viewBox="0 0 280 60" className={className} xmlns="http://www.w3.org/2000/svg">
    <text x="140" y="38" textAnchor="middle" fontFamily="'Cinzel', 'Copperplate', serif" fontWeight="700" fontSize="34" letterSpacing="6" fill={dark ? '#FFFEF2' : '#003553'}>AVANTE</text>
    <text x="140" y="55" textAnchor="middle" fontFamily="'Cinzel', 'Copperplate', serif" fontWeight="600" fontSize="9" letterSpacing="6" fill={dark ? '#FFFEF2' : '#003553'}>CAPE BRANDY</text>
  </svg>
);

const RaysBackdrop = ({ opacity = 0.07 }) => (
  <svg viewBox="0 0 400 400" className="absolute inset-0 w-full h-full pointer-events-none" preserveAspectRatio="xMidYMid slice" xmlns="http://www.w3.org/2000/svg">
    <g stroke="#FDB940" strokeWidth="0.4" opacity={opacity}>
      {Array.from({ length: 60 }).map((_, i) => {
        const angle = (i * 6) * Math.PI / 180;
        return <line key={i} x1="200" y1="200" x2={200 + 600 * Math.cos(angle)} y2={200 + 600 * Math.sin(angle)} />;
      })}
    </g>
  </svg>
);

// =================== Supabase helpers ===================
import { supabase } from './supabase.js';

// Map camelCase JS fields to snake_case DB columns and back
function clientToDb(c) {
  return {
    id: c.id,
    venue: c.venue || '',
    channel: c.channel || '',
    first_name: c.firstName || '',
    last_name: c.lastName || '',
    location: c.location || '',
    distributor: c.distributor || '',
    email: c.email || '',
    phone: c.phone || '',
    lead_source: c.leadSource || '',
    account_manager: c.accountManager || 'Unassigned',
    priority: c.priority || 'Medium',
    last_contacted: c.lastContacted || null,
    status: c.status || 'New',
    notes: c.notes || '',
    total_sales: c.totalSales || 0,
  };
}

function clientFromDb(r) {
  return {
    id: r.id,
    venue: r.venue || '',
    channel: r.channel || '',
    firstName: r.first_name || '',
    lastName: r.last_name || '',
    location: r.location || '',
    distributor: r.distributor || '',
    email: r.email || '',
    phone: r.phone || '',
    leadSource: r.lead_source || '',
    accountManager: r.account_manager || 'Unassigned',
    priority: r.priority || 'Medium',
    lastContacted: r.last_contacted || '',
    status: r.status || 'New',
    notes: r.notes || '',
    totalSales: Number(r.total_sales) || 0,
  };
}

function visitToDb(v) {
  return {
    client_id: v.clientId,
    sales_rep: v.salesRep || '',
    date: v.date || null,
    outcome: v.outcome || '',
    sale_amount: v.saleAmount || 0,
    items: v.items || [],
    notes: v.notes || '',
    follow_up_notes: v.followUpNotes || '',
    order_placed: v.orderPlaced || '',
  };
}

function visitFromDb(r) {
  return {
    id: r.id,
    clientId: r.client_id,
    salesRep: r.sales_rep || '',
    date: r.date || '',
    outcome: r.outcome || '',
    saleAmount: Number(r.sale_amount) || 0,
    items: r.items || [],
    notes: r.notes || '',
    followUpNotes: r.follow_up_notes || '',
    orderPlaced: r.order_placed || '',
  };
}

function targetToDb(rep, t) {
  return {
    rep,
    revenue: t.revenue || 0,
    visits: t.visits || 0,
    private_sales: t.privateSales || 0,
    trade_retail: t.tradeRetail || 0,
  };
}

function targetFromDb(r) {
  return {
    revenue: Number(r.revenue) || 0,
    visits: Number(r.visits) || 0,
    privateSales: Number(r.private_sales) || 0,
    tradeRetail: Number(r.trade_retail) || 0,
  };
}

// =================== Main App ===================
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { hasError: false, error: null }; }
  static getDerivedStateFromError(error) { return { hasError: true, error }; }
  componentDidCatch(error, info) { console.error('[ErrorBoundary]', error, info); }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: 40, fontFamily: 'Georgia, serif', color: '#003553', textAlign: 'center', minHeight: '100vh', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', background: '#FFFEF2' }}>
          <div style={{ fontFamily: 'Cinzel, serif', fontSize: 20, fontWeight: 700, letterSpacing: '0.15em', color: '#D78433' }}>AVANTE CRM</div>
          <p style={{ marginTop: 16, color: '#003553' }}>Something went wrong. Please reload.</p>
          <pre style={{ marginTop: 12, fontSize: 11, color: '#9c2c2c', background: '#f9f5e8', padding: 12, maxWidth: 500, overflow: 'auto', textAlign: 'left' }}>{this.state.error?.message || String(this.state.error)}</pre>
          <button onClick={() => window.location.reload()} style={{ marginTop: 16, padding: '10px 24px', background: '#D78433', color: '#fff', border: 'none', cursor: 'pointer', fontFamily: 'Cinzel, serif', letterSpacing: '0.15em', fontSize: 11 }}>RELOAD</button>
        </div>
      );
    }
    return this.props.children;
  }
}


export default function AvanteCRM() {
  const [view, setView] = useState('dashboard');
  const [loading, setLoading] = useState(true);
  const [clients, setClients] = useState([]);
  const [visits, setVisits] = useState([]);
  const [targets, setTargets] = useState(DEFAULT_TARGETS);
  const [activeRep, setActiveRep] = useState('All');
  const [showLogModal, setShowLogModal] = useState(false);
  const [editingVisit, setEditingVisit] = useState(null); // visit object being edited, or null
  const [selectedClient, setSelectedClient] = useState(null);
  const [newClientCtx, setNewClientCtx] = useState(null); // { defaultRep, onCreated } | null
  const [toast, setToast] = useState(null); // { message, type } | null
  const [confirmCtx, setConfirmCtx] = useState(null); // { title, message, onConfirm } | null

  const showToast = (message, type = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 3500);
  };

  // Custom in-app confirm — replaces window.confirm() which is blocked in
  // sandboxed iframes (claude.ai artifact runtime). Returns nothing — uses
  // callbacks instead, fired from the modal's Yes/No buttons.
  const askConfirm = ({ title, message, confirmLabel, danger, onConfirm }) => {
    setConfirmCtx({ title, message, confirmLabel, danger, onConfirm });
  };

  // Load all data from Supabase on mount
  useEffect(() => {
    (async () => {
      try {
        // Load clients
        const { data: clientRows, error: clientErr } = await supabase
          .from('clients').select('*').order('id');
        if (clientErr) throw clientErr;

        let loadedClients;
        if (!clientRows || clientRows.length === 0) {
          // First run — seed the database with the 278 venues
          const seedRows = SEED_CLIENTS.map(clientToDb);
          // Insert in batches of 50 to avoid request size limits
          for (let i = 0; i < seedRows.length; i += 50) {
            const { error } = await supabase.from('clients').insert(seedRows.slice(i, i + 50));
            if (error) console.error('[seed] clients batch error', error);
          }
          loadedClients = SEED_CLIENTS;
        } else {
          loadedClients = clientRows.map(clientFromDb);
        }

        // Load visits
        const { data: visitRows, error: visitErr } = await supabase
          .from('visits').select('*').order('created_at');
        if (visitErr) throw visitErr;

        // Load targets
        const { data: targetRows, error: targetErr } = await supabase
          .from('targets').select('*');
        if (targetErr) throw targetErr;

        let loadedTargets = { ...DEFAULT_TARGETS };
        if (!targetRows || targetRows.length === 0) {
          // Seed default targets
          const seedTargets = SALES_REPS.map(rep => targetToDb(rep, DEFAULT_TARGETS[rep]));
          const { error } = await supabase.from('targets').insert(seedTargets);
          if (error) console.error('[seed] targets error', error);
        } else {
          targetRows.forEach(r => {
            if (r.rep) loadedTargets[r.rep] = targetFromDb(r);
          });
        }

        setClients(loadedClients);
        setVisits((visitRows || []).map(visitFromDb));
        setTargets(loadedTargets);
      } catch (err) {
        console.error('[avante] init error', err);
        // Still load app with seed data so it's usable even if DB is down
        setClients(SEED_CLIENTS);
      }
      setLoading(false);
    })();
  }, []);

  // ------- Domain operations — write to Supabase AND update local state -------

  const persistTargets = async (next) => {
    setTargets(next);
    for (const rep of SALES_REPS) {
      if (next[rep]) {
        await supabase.from('targets')
          .upsert(targetToDb(rep, next[rep]), { onConflict: 'rep' });
      }
    }
  };

  const addVisit = async (visit) => {
    const computedTotal = (visit.items && visit.items.length > 0)
      ? visit.items.reduce((s, it) => s + (Number(it.unitPrice) || 0) * (Number(it.qty) || 0), 0)
      : (Number(visit.saleAmount) || 0);
    const enriched = { ...visit, saleAmount: computedTotal };

    // Insert visit into Supabase
    const { data: inserted, error } = await supabase
      .from('visits').insert(visitToDb(enriched)).select().single();
    if (error) { console.error('[addVisit]', error); return; }

    const newVisit = visitFromDb(inserted);
    setVisits((prev) => [...prev, newVisit]);

    // Update client total_sales in DB and local state
    const client = clients.find(c => c.id === visit.clientId);
    if (client) {
      const newSales = (client.totalSales || 0) + computedTotal;
      const newStatus = visit.outcome === 'Sold In' ? 'Converted'
        : visit.outcome === 'Rejected' ? 'Lost'
        : (client.status === 'New' ? 'Contacted' : client.status);
      await supabase.from('clients').update({
        total_sales: newSales,
        last_contacted: visit.date,
        status: newStatus,
      }).eq('id', visit.clientId);
      setClients((prev) => prev.map(c => c.id === visit.clientId
        ? { ...c, totalSales: newSales, lastContacted: visit.date, status: newStatus }
        : c));
    }
    return newVisit;
  };

  const updateVisit = async (visitId, patch) => {
    const oldVisit = visits.find(v => v.id === visitId);
    if (!oldVisit) return;
    const merged = { ...oldVisit, ...patch };
    const newTotal = (merged.items && merged.items.length > 0)
      ? merged.items.reduce((s, it) => s + (Number(it.unitPrice) || 0) * (Number(it.qty) || 0), 0)
      : 0;
    merged.saleAmount = newTotal;

    const { error } = await supabase.from('visits')
      .update(visitToDb(merged)).eq('id', visitId);
    if (error) { console.error('[updateVisit]', error); return; }

    setVisits((prev) => prev.map(v => v.id === visitId ? merged : v));

    const oldTotal = Number(oldVisit.saleAmount) || 0;
    const oldClientId = oldVisit.clientId;
    const newClientId = merged.clientId;

    setClients((prev) => prev.map(c => {
      if (oldClientId === newClientId && c.id === oldClientId) {
        const updated = { ...c, totalSales: Math.max(0, (c.totalSales || 0) + (newTotal - oldTotal)), lastContacted: merged.date };
        supabase.from('clients').update({ total_sales: updated.totalSales, last_contacted: updated.lastContacted }).eq('id', c.id);
        return updated;
      }
      if (c.id === oldClientId && oldClientId !== newClientId) {
        const updated = { ...c, totalSales: Math.max(0, (c.totalSales || 0) - oldTotal) };
        supabase.from('clients').update({ total_sales: updated.totalSales }).eq('id', c.id);
        return updated;
      }
      if (c.id === newClientId && oldClientId !== newClientId) {
        const updated = { ...c, totalSales: (c.totalSales || 0) + newTotal, lastContacted: merged.date };
        supabase.from('clients').update({ total_sales: updated.totalSales, last_contacted: updated.lastContacted }).eq('id', c.id);
        return updated;
      }
      return c;
    }));
  };

  const deleteVisit = async (visitId) => {
    const v = visits.find(x => x.id === visitId);
    if (!v) return;
    const refund = Number(v.saleAmount) || 0;
    const clientId = v.clientId;

    const { error } = await supabase.from('visits').delete().eq('id', visitId);
    if (error) { console.error('[deleteVisit]', error); return; }

    setVisits((prev) => prev.filter(x => x.id !== visitId));
    setClients((prev) => prev.map(c => {
      if (c.id === clientId) {
        const updated = { ...c, totalSales: Math.max(0, (c.totalSales || 0) - refund) };
        supabase.from('clients').update({ total_sales: updated.totalSales }).eq('id', c.id);
        return updated;
      }
      return c;
    }));
  };

  const updateClient = async (id, patch) => {
    const dbPatch = {};
    if (patch.venue !== undefined) dbPatch.venue = patch.venue;
    if (patch.firstName !== undefined) dbPatch.first_name = patch.firstName;
    if (patch.lastName !== undefined) dbPatch.last_name = patch.lastName;
    if (patch.email !== undefined) dbPatch.email = patch.email;
    if (patch.phone !== undefined) dbPatch.phone = patch.phone;
    if (patch.location !== undefined) dbPatch.location = patch.location;
    if (patch.distributor !== undefined) dbPatch.distributor = patch.distributor;
    if (patch.leadSource !== undefined) dbPatch.lead_source = patch.leadSource;
    if (patch.accountManager !== undefined) dbPatch.account_manager = patch.accountManager;
    if (patch.priority !== undefined) dbPatch.priority = patch.priority;
    if (patch.status !== undefined) dbPatch.status = patch.status;
    if (patch.notes !== undefined) dbPatch.notes = patch.notes;
    if (patch.channel !== undefined) dbPatch.channel = patch.channel;
    if (patch.totalSales !== undefined) dbPatch.total_sales = patch.totalSales;
    if (patch.lastContacted !== undefined) dbPatch.last_contacted = patch.lastContacted;
    if (Object.keys(dbPatch).length > 0) {
      await supabase.from('clients').update(dbPatch).eq('id', id);
    }
    setClients((prev) => prev.map(c => c.id === id ? { ...c, ...patch } : c));
  };

  const clientIdCounter = useRef(null);
  const addClient = async (client) => {
    const newClientData = {
      venue: '', firstName: '', lastName: '', location: '', distributor: '',
      email: '', phone: '', leadSource: 'Cold Call', accountManager: 'Unassigned',
      priority: 'Medium', lastContacted: '', status: 'New', notes: '',
      totalSales: 0, channel: '',
      ...client,
    };
    const dbData = clientToDb({ id: undefined, ...newClientData });
    console.log('[addClient] inserting:', dbData);
    const { data: inserted, error } = await supabase
      .from('clients').insert(dbData)
      .select().single();
    if (error) {
      console.error('[addClient] error:', error);
      console.error('[addClient] error details:', JSON.stringify(error));
      throw new Error(error.message || JSON.stringify(error));
    }
    const newClient = clientFromDb(inserted);
    setClients((prev) => [...prev, newClient]);
    return newClient;
  };

  const monthVisits = useMemo(() => {
    const m = monthISO();
    return visits.filter(v => v.date && v.date.startsWith(m));
  }, [visits]);

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: '#003553' }}>
        <div className="text-center">
          <div className="inline-block animate-pulse">
            <AvanteLogo height={140} />
          </div>
          <p className="mt-6 text-sm tracking-[0.4em]" style={{ color: '#FDB940', fontFamily: "'Cinzel', serif", fontWeight: 600 }}>LOADING CRM</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen overflow-x-hidden" style={{ background: '#FFFEF2', fontFamily: "'Libre Baskerville', Georgia, serif", color: '#003553' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Libre+Baskerville:ital,wght@0,400;0,700;1,400&family=Cinzel:wght@500;600;700;900&display=swap');
        *, *::before, *::after { box-sizing: border-box; }
        html, body { margin: 0; padding: 0; overflow-x: hidden; width: 100%; }
        .font-display { font-family: 'Cinzel', 'Copperplate', serif; letter-spacing: 0.08em; }
        .font-body { font-family: 'Libre Baskerville', Georgia, serif; }
        /* Colour utilities */
        .ink { color: #003553; }
        .ocean { color: #006C90; }
        .gold { color: #FDB940; }
        .copper { color: #D78433; }
        .bg-ink { background: #003553; }
        .bg-cream { background: #FFFEF2; }
        .bg-gold { background: #FDB940; }
        .bg-copper { background: #D78433; }
        .border-ink { border-color: #003553; }
        .border-copper { border-color: #D78433; }
        .border-gold { border-color: #FDB940; }
        .diamond-clip { clip-path: polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%); }
        .premium-card { background: #FFFEF2; border: 1px solid rgba(0,53,83,0.15); box-shadow: 0 1px 0 rgba(0,53,83,0.04), 0 8px 24px -16px rgba(0,53,83,0.2); }
        /* Scrollbar */
        .scrollbar-thin::-webkit-scrollbar { width: 4px; height: 4px; }
        .scrollbar-thin::-webkit-scrollbar-track { background: rgba(0,53,83,0.05); }
        .scrollbar-thin::-webkit-scrollbar-thumb { background: rgba(0,53,83,0.3); border-radius: 3px; }
        .no-scrollbar::-webkit-scrollbar { display: none; }
        .no-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }
        /* Animations */
        @keyframes fadeUp { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
        .fade-up { animation: fadeUp 0.4s ease-out forwards; }
        @keyframes slideUp { from { transform: translateY(100%); opacity: 0.8; } to { transform: translateY(0); opacity: 1; } }
        .modal-slide-up { animation: slideUp 0.25s ease-out forwards; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .animate-spin { animation: spin 1s linear infinite; }
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.5} }
        .animate-pulse { animation: pulse 2s cubic-bezier(.4,0,.6,1) infinite; }
        /* ── HEADER ── */
        .crm-header {
          background: #003553;
          border-bottom: 2px solid rgba(255,219,64,0.25);
          width: 100%;
          position: sticky;
          top: 0;
          z-index: 50;
        }
        .crm-header-inner {
          display: flex;
          align-items: center;
          width: 100%;
          max-width: 1400px;
          margin: 0 auto;
          padding: 0 16px;
          height: 56px;
          gap: 12px;
        }
        .crm-logo-wrap {
          display: flex;
          align-items: center;
          gap: 10px;
          flex-shrink: 0;
          text-decoration: none;
        }
        .crm-brand-text { display: none; }
        .crm-nav {
          display: flex;
          align-items: center;
          gap: 2px;
          flex: 1;
          overflow-x: auto;
          -ms-overflow-style: none;
          scrollbar-width: none;
        }
        .crm-nav::-webkit-scrollbar { display: none; }
        .crm-nav-btn {
          display: flex;
          align-items: center;
          gap: 5px;
          padding: 7px 10px;
          font-family: 'Cinzel', 'Copperplate', serif;
          font-size: 9px;
          font-weight: 600;
          letter-spacing: 0.15em;
          white-space: nowrap;
          cursor: pointer;
          border: none;
          background: transparent;
          color: rgba(255,255,255,0.65);
          transition: color 0.15s, background 0.15s;
          flex-shrink: 0;
        }
        .crm-nav-btn:hover { color: #fff; }
        .crm-nav-btn.active { background: #FDB940; color: #003553; }
        .crm-log-btn {
          flex-shrink: 0;
          display: flex;
          align-items: center;
          gap: 6px;
          padding: 8px 14px;
          font-family: 'Cinzel', 'Copperplate', serif;
          font-size: 9px;
          font-weight: 700;
          letter-spacing: 0.2em;
          background: #D78433;
          color: #FFFEF2;
          border: none;
          cursor: pointer;
          white-space: nowrap;
        }
        .crm-log-btn:hover { background: #FDB940; color: #003553; }
        /* Show brand text on wider screens */
        @media (min-width: 640px) {
          .crm-header-inner { padding: 0 24px; height: 60px; }
          .crm-brand-text { display: block; }
          .crm-nav-btn { font-size: 10px; padding: 8px 12px; }
          .crm-log-btn { font-size: 10px; padding: 9px 18px; }
        }
        @media (min-width: 1024px) {
          .crm-header-inner { padding: 0 40px; }
        }
        /* ── MAIN CONTENT ── */
        .crm-main {
          max-width: 1400px;
          margin: 0 auto;
          padding: 16px 16px 80px;
        }
        @media (min-width: 640px) {
          .crm-main { padding: 24px 24px 32px; }
        }
        @media (min-width: 1024px) {
          .crm-main { padding: 32px 40px; }
        }
        /* ── GRID HELPERS (replace Tailwind grid) ── */
        .grid-1 { display: grid; grid-template-columns: 1fr; gap: 12px; }
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
        .grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; }
        .grid-6 { display: grid; grid-template-columns: repeat(3,1fr); gap: 8px; }
        @media (min-width: 640px) {
          .grid-kpi { display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; }
          .grid-6 { grid-template-columns: repeat(6,1fr); }
          .grid-lb { display: grid; grid-template-columns: 2fr 1fr; gap: 16px; }
          .grid-repcard { display: grid; grid-template-columns: repeat(3,1fr); gap: 16px; }
        }
        .grid-kpi { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
        .grid-lb { display: grid; grid-template-columns: 1fr; gap: 16px; }
        .grid-repcard { display: grid; grid-template-columns: 1fr; gap: 16px; }
        /* ── SPACING / FLEX helpers ── */
        .flex { display: flex; }
        .flex-col { flex-direction: column; }
        .flex-wrap { flex-wrap: wrap; }
        .items-center { align-items: center; }
        .items-start { align-items: flex-start; }
        .justify-between { justify-content: space-between; }
        .justify-center { justify-content: center; }
        .justify-end { justify-content: flex-end; }
        .gap-1 { gap: 4px; }
        .gap-2 { gap: 8px; }
        .gap-3 { gap: 12px; }
        .gap-4 { gap: 16px; }
        .gap-5 { gap: 20px; }
        .flex-1 { flex: 1; }
        .flex-shrink-0 { flex-shrink: 0; }
        .min-w-0 { min-width: 0; }
        .w-full { width: 100%; }
        .text-center { text-align: center; }
        .text-left { text-align: left; }
        .relative { position: relative; }
        .absolute { position: absolute; }
        .fixed { position: fixed; }
        .sticky { position: sticky; top: 0; }
        .inset-0 { top:0; right:0; bottom:0; left:0; position:absolute; }
        .overflow-hidden { overflow: hidden; }
        .overflow-x-auto { overflow-x: auto; }
        .overflow-y-auto { overflow-y: auto; }
        .overflow-x-hidden { overflow-x: hidden; }
        .truncate { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .italic { font-style: italic; }
        .block { display: block; }
        .inline-block { display: inline-block; }
        .hidden { display: none; }
        .pointer-events-none { pointer-events: none; }
        .cursor-pointer { cursor: pointer; }
        .cursor-not-allowed { cursor: not-allowed; }
        .select-none { user-select: none; }
        .transition-all { transition: all 0.15s ease; }
        .transition-colors { transition: color 0.15s ease, background-color 0.15s ease; }
        /* Padding / Margin helpers */
        .p-2 { padding: 8px; } .p-3 { padding: 12px; } .p-4 { padding: 16px; }
        .p-5 { padding: 20px; } .p-6 { padding: 24px; } .p-10 { padding: 40px; }
        .px-2 { padding-left:8px; padding-right:8px; }
        .px-3 { padding-left:12px; padding-right:12px; }
        .px-4 { padding-left:16px; padding-right:16px; }
        .px-5 { padding-left:20px; padding-right:20px; }
        .py-1 { padding-top:4px; padding-bottom:4px; }
        .py-2 { padding-top:8px; padding-bottom:8px; }
        .py-2\.5 { padding-top:10px; padding-bottom:10px; }
        .py-3 { padding-top:12px; padding-bottom:12px; }
        .py-4 { padding-top:16px; padding-bottom:16px; }
        .pb-3 { padding-bottom:12px; }
        .pb-4 { padding-bottom:16px; }
        .pt-3 { padding-top:12px; }
        .pt-4 { padding-top:16px; }
        .mt-0\.5 { margin-top:2px; } .mt-1 { margin-top:4px; } .mt-2 { margin-top:8px; }
        .mt-3 { margin-top:12px; } .mt-4 { margin-top:16px; } .mt-6 { margin-top:24px; }
        .mb-1 { margin-bottom:4px; } .mb-2 { margin-bottom:8px; } .mb-3 { margin-bottom:12px; }
        .mb-4 { margin-bottom:16px; } .mb-5 { margin-bottom:20px; }
        .ml-auto { margin-left:auto; }
        .mx-auto { margin-left:auto; margin-right:auto; }
        .space-y-2 > * + * { margin-top: 8px; }
        .space-y-3 > * + * { margin-top: 12px; }
        .space-y-4 > * + * { margin-top: 16px; }
        .space-y-5 > * + * { margin-top: 20px; }
        .space-y-6 > * + * { margin-top: 24px; }
        /* Border */
        .border { border: 1px solid rgba(0,53,83,0.2); }
        .border-2 { border: 2px solid; }
        .border-b { border-bottom: 1px solid rgba(0,53,83,0.15); }
        .border-t { border-top: 1px solid rgba(0,53,83,0.15); }
        .border-l-4 { border-left: 4px solid; }
        .border-l-2 { border-left: 2px solid; }
        .rounded { border-radius: 4px; }
        .rounded-full { border-radius: 9999px; }
        /* Sizing */
        .w-2 { width:8px; } .w-3 { width:12px; } .w-3\.5 { width:14px; }
        .w-4 { width:16px; } .w-5 { width:20px; } .w-6 { width:24px; }
        .w-8 { width:32px; } .w-10 { width:40px; } .w-16 { width:64px; }
        .h-1 { height:4px; } .h-1\.5 { height:6px; } .h-2 { height:8px; }
        .h-3 { height:12px; } .h-3\.5 { height:14px; } .h-4 { height:16px; }
        .h-5 { height:20px; } .h-6 { height:24px; } .h-8 { height:32px; }
        .h-10 { height:40px; } .h-full { height:100%; }
        .min-h-screen { min-height: 100vh; }
        .max-h-\[95vh\] { max-height: 95vh; }
        .max-h-\[80vh\] { max-height: 80vh; }
        .max-h-\[60vh\] { max-height: 60vh; }
        .max-w-2xl { max-width: 672px; }
        .max-w-3xl { max-width: 768px; }
        /* Text sizes */
        .text-xs { font-size: 12px; } .text-sm { font-size: 14px; }
        .text-base { font-size: 16px; } .text-lg { font-size: 18px; }
        .text-xl { font-size: 20px; } .text-2xl { font-size: 24px; }
        .text-3xl { font-size: 30px; } .text-4xl { font-size: 36px; }
        /* Font weights */
        .font-bold { font-weight: 700; }
        /* Opacity */
        .opacity-40 { opacity: 0.4; } .opacity-60 { opacity: 0.6; }
        /* Z-index */
        .z-10 { z-index: 10; } .z-30 { z-index: 30; } .z-40 { z-index: 40; }
        .z-50 { z-index: 50; } .z-60 { z-index: 60; }
        /* Shadow */
        .shadow-2xl { box-shadow: 0 25px 50px -12px rgba(0,0,0,.25); }
        /* Backdrop */
        .backdrop-blur-sm { backdrop-filter: blur(4px); }
        /* Resize */
        .resize-none { resize: none; }
        /* Focus */
        input:focus, select:focus, textarea:focus { outline: none; border-color: #D78433 !important; }
        /* Mobile form inputs — prevent zoom on iOS */
        @media (max-width: 639px) {
          input, select, textarea { font-size: 16px !important; }
        }
        /* Hover states */
        .hover-bg:hover { background: rgba(0,53,83,0.06); }
      `}</style>

      <Header view={view} setView={setView} onLog={() => setShowLogModal(true)} />

      <main className="crm-main">
        {view === 'dashboard' && (
          <Dashboard
            clients={clients}
            visits={monthVisits}
            allVisits={visits}
            targets={targets}
            activeRep={activeRep}
            setActiveRep={setActiveRep}
            onNavigate={setView}
          />
        )}
        {view === 'leads' && (
          <LeadsPage
            clients={clients}
            visits={visits}
            updateClient={updateClient}
            onSelect={setSelectedClient}
          />
        )}
        {view === 'visits' && (
          <VisitsPage
            visits={visits}
            clients={clients}
            onLog={() => setShowLogModal(true)}
            onEdit={(visit) => setEditingVisit(visit)}
            onDelete={(visitId) => {
              const v = visits.find(x => x.id === visitId);
              const c = v ? clients.find(cl => cl.id === v.clientId) : null;
              const saleNote = v?.saleAmount > 0
                ? `Sale of R ${Math.round(v.saleAmount).toLocaleString('en-ZA')} will be removed from the client total.`
                : '';
              askConfirm({
                title: 'Delete this visit?',
                message: `${c?.venue || 'Unknown'} · ${v?.date || ''}\nOutcome: ${v?.outcome || ''}${saleNote ? '\n' + saleNote : ''}\n\nThis cannot be undone.`,
                confirmLabel: 'DELETE VISIT',
                danger: true,
                onConfirm: () => {
                  deleteVisit(visitId);
                  showToast('Visit deleted · client totals updated', 'success');
                },
              });
            }}
          />
        )}
        {view === 'manager' && (
          <ManagerPortal
            targets={targets}
            saveTargets={persistTargets}
            clients={clients}
            visits={visits}
            askConfirm={askConfirm}
          />
        )}
      </main>

      {showLogModal && (
        <LogVisitModal
          clients={clients}
          onClose={() => setShowLogModal(false)}
          onSubmit={async (v) => {
            await addVisit(v);
            setShowLogModal(false);
            const c = clients.find(cl => cl.id === v.clientId);
            const total = (v.items || []).reduce((s, it) => s + (Number(it.unitPrice) || 0) * (Number(it.qty) || 0), 0);
            showToast(`Visit saved: ${c?.venue || 'venue'} · ${total > 0 ? 'R ' + Math.round(total).toLocaleString('en-ZA') : v.outcome}`);
          }}
          onRequestNewClient={(defaultRep, onCreated) => setNewClientCtx({ defaultRep, onCreated })}
        />
      )}

      {editingVisit && (
        <LogVisitModal
          clients={clients}
          existingVisit={editingVisit}
          onClose={() => setEditingVisit(null)}
          onSubmit={async (v) => {
            await updateVisit(editingVisit.id, v);
            setEditingVisit(null);
            const c = clients.find(cl => cl.id === v.clientId);
            const total = (v.items || []).reduce((s, it) => s + (Number(it.unitPrice) || 0) * (Number(it.qty) || 0), 0);
            showToast(`Visit updated: ${c?.venue || 'venue'} · ${total > 0 ? 'R ' + Math.round(total).toLocaleString('en-ZA') : v.outcome}`);
          }}
          onRequestNewClient={(defaultRep, onCreated) => setNewClientCtx({ defaultRep, onCreated })}
        />
      )}

      {selectedClient && (
        <ClientDetailModal
          client={selectedClient}
          visits={visits.filter(v => v.clientId === selectedClient.id)}
          onClose={() => setSelectedClient(null)}
          onUpdate={(patch) => {
            updateClient(selectedClient.id, patch);
            setSelectedClient({ ...selectedClient, ...patch });
            showToast(`Client updated: ${patch.venue || selectedClient.venue}`);
          }}
        />
      )}

      {newClientCtx && (
        <NewClientModal
          defaultRep={newClientCtx.defaultRep}
          onClose={() => setNewClientCtx(null)}
          onSave={async (data) => {
            const created = await addClient(data);
            if (!created) return null; // modal will show error
            if (newClientCtx.onCreated) newClientCtx.onCreated(created);
            setNewClientCtx(null);
            showToast(`✓ New client saved: ${created.venue} → ${created.accountManager}'s book`);
            return created;
          }}
        />
      )}

      {toast && (
        <div className="fixed bottom-20 md:bottom-6 right-4 md:right-6 z-[80]">
          <div className="flex items-center gap-3 px-4 py-3 shadow-2xl border-l-4" style={{ background: '#003553', borderLeftColor: toast.type === 'error' ? '#9c2c2c' : '#FDB940', minWidth: '240px', maxWidth: '320px' }}>
            <div className="w-2 h-2 flex-shrink-0 diamond-clip" style={{ background: toast.type === 'error' ? '#9c2c2c' : '#FDB940' }}></div>
            <div className="min-w-0">
              <p className="font-display text-[9px] tracking-[0.3em]" style={{ color: '#FDB940', fontWeight: 600 }}>{toast.type === 'error' ? 'ERROR' : 'SAVED'}</p>
              <p className="text-xs truncate" style={{ color: '#FFFEF2', fontWeight: 500 }}>{toast.message}</p>
            </div>
          </div>
        </div>
      )}

      {confirmCtx && (
        <ConfirmModal
          title={confirmCtx.title}
          message={confirmCtx.message}
          confirmLabel={confirmCtx.confirmLabel}
          danger={confirmCtx.danger}
          onCancel={() => setConfirmCtx(null)}
          onConfirm={() => {
            const fn = confirmCtx.onConfirm;
            setConfirmCtx(null);
            if (fn) fn();
          }}
        />
      )}
    </div>
  );
}

// =================== Confirm Modal ===================
function ConfirmModal({ title, message, confirmLabel, danger, onCancel, onConfirm }) {
  return (
    <div style={{ position:"fixed", inset:0, zIndex:90, display:"flex", alignItems:"center", justifyContent:"center", padding:16, background:"rgba(0,53,83,0.82)", backdropFilter:"blur(4px)" }} onClick={onCancel}>
      <div className="bg-cream max-w-md w-full border-2 border-copper shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="p-5 flex items-center gap-3" style={{ background: '#003553' }}>
          <div className="w-2.5 h-2.5 diamond-clip" style={{ background: danger ? '#9c2c2c' : '#FDB940' }}></div>
          <div>
            <p className="font-display text-[10px] tracking-[0.4em]" style={{ color: '#FDB940', fontWeight: 600 }}>{danger ? 'DESTRUCTIVE ACTION' : 'CONFIRM'}</p>
            <h2 className="font-display text-lg mt-0.5" style={{ color: '#FFFEF2', fontWeight: 700, letterSpacing: '0.06em' }}>{title}</h2>
          </div>
        </div>
        <div className="p-6">
          <p className="text-sm ink whitespace-pre-line">{message}</p>
        </div>
        <div className="flex justify-end gap-3 px-6 pb-6">
          <button
            type="button"
            onClick={onCancel}
            className="px-5 py-2.5 font-display text-xs tracking-[0.25em] ink hover:"
            style={{ fontWeight: 700 }}
          >
            CANCEL
          </button>
          <button
            type="button"
            onClick={onConfirm}
            className="px-6 py-2.5 font-display text-xs tracking-[0.25em] flex items-center gap-2 hover:opacity-90"
            style={{
              background: danger ? '#9c2c2c' : '#003553',
              color: '#FFFEF2',
              fontWeight: 700,
            }}
            autoFocus
          >
            {confirmLabel || 'CONFIRM'}
          </button>
        </div>
      </div>
    </div>
  );
}

// =================== Header ===================
function Header({ view, setView, onLog }) {
  const tabs = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'leads', label: 'Leads & Clients', icon: Users },
    { id: 'visits', label: 'Visit Log', icon: ClipboardList },
    { id: 'manager', label: 'Manager', icon: Settings },
  ];
  return (
    <header className="crm-header">
      <div className="crm-header-inner">
        {/* Logo */}
        <div className="crm-logo-wrap">
          <AvanteLogo height={40} />
          <div className="crm-brand-text" style={{ borderLeft: '1px solid rgba(255,255,255,0.2)', paddingLeft: 12 }}>
            <p className="font-display" style={{ fontSize: 8, letterSpacing: '0.35em', color: '#FDB940', fontWeight: 600, margin: 0 }}>SALES CRM</p>
            <p className="font-display" style={{ fontSize: 7, letterSpacing: '0.2em', color: '#FFFEF2', opacity: 0.5, margin: '2px 0 0' }}>DARE TO FORWARD</p>
          </div>
        </div>
        {/* Navigation — horizontal tabs, always visible */}
        <nav className="crm-nav">
          {tabs.map(t => {
            const Icon = t.icon;
            return (
              <button key={t.id} onClick={() => setView(t.id)}
                className={'crm-nav-btn' + (view === t.id ? ' active' : '')}>
                <Icon style={{ width: 13, height: 13, flexShrink: 0 }} />
                <span>{t.label.toUpperCase()}</span>
              </button>
            );
          })}
        </nav>
        {/* Log Visit CTA */}
        <button onClick={onLog} className="crm-log-btn">
          <Plus style={{ width: 14, height: 14 }} />
          <span>LOG VISIT</span>
        </button>
      </div>
    </header>
  );
}

// =================== Dashboard ===================
function Dashboard({ clients, visits, allVisits, targets, activeRep, setActiveRep, onNavigate }) {
  const repFilter = (rec, key = 'accountManager') => activeRep === 'All' || rec[key] === activeRep;
  const filteredClients = clients.filter(c => repFilter(c));
  const filteredVisits = visits.filter(v => repFilter(v, 'salesRep'));

  const totalSales = filteredClients.reduce((s, c) => s + (c.totalSales || 0), 0);
  const monthSales = filteredVisits.reduce((s, v) => s + (v.saleAmount || 0), 0);
  const visitCount = filteredVisits.length;
  const sold = filteredVisits.filter(v => v.outcome === 'Sold In').length;
  const conversionRate = visitCount > 0 ? Math.round((sold / visitCount) * 100) : 0;

  const activeTargets = useMemo(() => {
    if (activeRep === 'All') {
      return SALES_REPS.reduce((acc, r) => ({
        revenue: acc.revenue + (targets[r]?.revenue || 0),
        visits: acc.visits + (targets[r]?.visits || 0),
        privateSales: acc.privateSales + (targets[r]?.privateSales || 0),
        tradeRetail: acc.tradeRetail + (targets[r]?.tradeRetail || 0),
      }), { revenue: 0, visits: 0, privateSales: 0, tradeRetail: 0 });
    }
    return targets[activeRep] || { revenue: 0, visits: 0, privateSales: 0, tradeRetail: 0 };
  }, [activeRep, targets]);

  const pct = (a, b) => b > 0 ? Math.min(100, Math.round((a / b) * 100)) : 0;

  // Stage counts for pipeline summary
  const stages = ['New', 'Contacted', 'Qualified', 'Prospect', 'Converted', 'Lost'];
  const stageCounts = Object.fromEntries(stages.map(s => [s, filteredClients.filter(c => c.status === s).length]));
  const stageColors = { New: '#006C90', Contacted: '#FDB940', Qualified: '#D78433', Prospect: '#003553', Converted: '#2d8659', Lost: '#9c2c2c' };

  return (
    <div className="space-y-4 md:space-y-6 fade-up">

      {/* Page header */}
      <div className="flex items-center justify-between gap-3 flex-wrap pb-3 border-b border">
        <div>
          <p className="font-display text-[9px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>PERFORMANCE OVERVIEW</p>
          <h1 className="font-display ink mt-0.5" style={{ fontWeight: 700, fontSize: 28 }}>
            {new Date().toLocaleDateString('en-ZA', { month: 'long', year: 'numeric' }).toUpperCase()}
          </h1>
        </div>
        <RepToggle active={activeRep} onChange={setActiveRep} />
      </div>

      {/* ── KPI CARDS ── */}
      <div className="grid-kpi">
        <KPICard label="Month Sales" value={ZAR(monthSales)} target={activeTargets.revenue} targetValue={ZAR(activeTargets.revenue)} progress={pct(monthSales, activeTargets.revenue)} icon={DollarSign} accent="copper" />
        <KPICard label="Visits" value={visitCount} target={activeTargets.visits} targetValue={`${activeTargets.visits} visits`} progress={pct(visitCount, activeTargets.visits)} icon={Activity} accent="ocean" />
        <KPICard label="Conversion" value={`${conversionRate}%`} subtitle={`${sold} of ${visitCount} sold`} progress={conversionRate} icon={TrendingUp} accent="gold" />
        <KPICard label="Total TD" value={ZAR(totalSales)} subtitle="All-time" icon={Award} accent="ink" />
      </div>

      {/* ── LEADERBOARD + RECENT VISITS ── */}
      <div className="grid-lb">
        <div className="premium-card p-4">
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center gap-2">
              <div className="w-2 h-2 bg-copper diamond-clip"></div>
              <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>LEADERBOARD</h2>
            </div>
            <span className="text-[9px] tracking-[0.15em] font-display ocean uppercase">This month</span>
          </div>
          <Leaderboard clients={clients} visits={allVisits} targets={targets} />
        </div>

        <div className="premium-card p-4">
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center gap-2">
              <div className="w-2 h-2 bg-copper diamond-clip"></div>
              <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>RECENT VISITS</h2>
            </div>
            <button onClick={() => onNavigate('visits')}
              className="text-[9px] font-display tracking-[0.15em] copper hover:opacity-70 flex items-center gap-1" style={{ fontWeight: 700 }}>
              ALL <ChevronRight className="w-3 h-3" />
            </button>
          </div>
          <RecentVisits visits={filteredVisits.slice(-6).reverse()} clients={clients} />
        </div>
      </div>

      {/* ── PIPELINE SUMMARY (compact) ── */}
      <div className="premium-card p-4">
        <div className="flex items-center justify-between mb-3">
          <div className="flex items-center gap-2">
            <div className="w-2 h-2 bg-copper diamond-clip"></div>
            <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>
              PIPELINE {activeRep !== 'All' && `— ${activeRep.toUpperCase()}`}
            </h2>
          </div>
          <button onClick={() => onNavigate('leads')}
            className="text-[9px] font-display tracking-[0.15em] copper hover:opacity-70 flex items-center gap-1" style={{ fontWeight: 700 }}>
            VIEW ALL <ChevronRight className="w-3 h-3" />
          </button>
        </div>
        <div className="grid-6">
          {stages.map(s => (
            <button key={s} onClick={() => onNavigate('leads')}
              className="text-center p-2 hover: transition-colors cursor-pointer">
              <div className="font-display text-xl md:text-2xl mb-0.5" style={{ fontWeight: 700, color: stageColors[s] }}>{stageCounts[s]}</div>
              <div className="font-display text-[8px] md:text-[9px] tracking-[0.1em] ocean mb-1.5" style={{ fontWeight: 600 }}>{s.toUpperCase()}</div>
              <div className="h-1 ">
                <div className="h-full" style={{ width: `${Math.max(...Object.values(stageCounts), 1) > 0 ? (stageCounts[s] / Math.max(...Object.values(stageCounts), 1)) * 100 : 0}%`, background: stageColors[s] }}></div>
              </div>
            </button>
          ))}
        </div>
      </div>

      {/* ── QUICK NAV CARDS — tap to go to each section ── */}
      <div className="grid-3">
        {[
          { id: 'leads', label: 'Leads & Clients', sub: `${filteredClients.length} venues`, icon: Users, color: '#006C90' },
          { id: 'visits', label: 'Visit Log', sub: `${filteredVisits.length} this month`, icon: ClipboardList, color: '#D78433' },
          { id: 'manager', label: 'Manager Portal', sub: 'Targets & Export', icon: Settings, color: '#003553' },
        ].map(item => {
          const Icon = item.icon;
          return (
            <button key={item.id} onClick={() => onNavigate(item.id)}
              className="premium-card p-4 text-left hover:shadow-md transition-all active:scale-[0.98] flex items-center gap-3">
              <div className="w-10 h-10 flex-shrink-0 flex items-center justify-center" style={{ background: item.color }}>
                <Icon className="w-5 h-5" style={{ color: '#FFFEF2' }} />
              </div>
              <div className="min-w-0">
                <p className="font-display text-xs tracking-[0.1em] ink truncate" style={{ fontWeight: 700 }}>{item.label.toUpperCase()}</p>
                <p className="text-[10px] italic ocean mt-0.5">{item.sub}</p>
              </div>
              <ChevronRight className="w-4 h-4 copper flex-shrink-0 ml-auto" />
            </button>
          );
        })}
      </div>

    </div>
  );
}

function RepToggle({ active, onChange }) {
  const opts = ['All', ...SALES_REPS];
  return (
    <div className="inline-flex border border bg-cream">
      {opts.map((r, i) => {
        const isActive = active === r;
        return (
          <button
            key={r}
            onClick={() => onChange(r)}
            className={`px-2.5 md:px-4 py-1.5 md:py-2 font-display text-[9px] md:text-[11px] tracking-[0.2em] transition-all ${isActive ? 'bg-ink' : 'hover:'} ${i > 0 ? 'border-l border' : ''}`}
            style={{ fontWeight: 600, color: isActive ? '#FFFEF2' : '#003553' }}
          >
            {r === 'All' ? 'ALL' : r.toUpperCase()}
          </button>
        );
      })}
    </div>
  );
}

function KPICard({ label, value, subtitle, target, targetValue, progress, icon: Icon, accent }) {
  const accentMap = {
    copper: { bg: '#D78433', text: '#D78433' },
    ocean: { bg: '#006C90', text: '#006C90' },
    gold: { bg: '#FDB940', text: '#D78433' },
    ink: { bg: '#003553', text: '#003553' },
  };
  const a = accentMap[accent] || accentMap.ink;
  return (
    <div className="premium-card p-3 md:p-5 relative overflow-hidden">
      <div className="absolute top-0 right-0 w-16 h-16 opacity-[0.04]" style={{ background: a.bg, clipPath: 'polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)', transform: 'translate(30%, -30%)' }}></div>
      <div className="flex items-center justify-between mb-2 relative">
        <span className="font-display text-[9px] tracking-[0.2em]" style={{ color: a.text, fontWeight: 600 }}>{label.toUpperCase()}</span>
        <Icon className="w-3.5 h-3.5" style={{ color: a.text }} />
      </div>
      <div className="font-display text-2xl md:text-3xl ink mb-1" style={{ fontWeight: 700 }}>{value}</div>
      {target !== undefined && progress !== undefined ? (
        <>
          <div className="flex items-center justify-between text-[10px] mt-2 mb-1">
            <span className="italic ocean hidden md:inline">Target {targetValue}</span>
            <span className="italic ocean md:hidden">{progress}%</span>
            <span className="font-display tracking-wider hidden md:inline" style={{ color: a.text, fontWeight: 700 }}>{progress}%</span>
          </div>
          <div className="h-1  relative overflow-hidden">
            <div className="h-full transition-all duration-700" style={{ width: `${progress}%`, background: a.bg }}></div>
          </div>
        </>
      ) : (
        <p className="text-[10px] italic ocean mt-1">{subtitle}</p>
      )}
    </div>
  );
}

function ChannelCard({ title, subtitle, icon: Icon, visits, target }) {
  const pct = target > 0 ? Math.min(100, Math.round((visits / target) * 100)) : 0;
  return (
    <div className="premium-card p-5">
      <div className="flex items-start gap-4">
        <div className="w-12 h-12 flex-shrink-0 flex items-center justify-center  border border">
          <Icon className="w-5 h-5 ink" />
        </div>
        <div className="flex-1 min-w-0">
          <h3 className="font-display text-base ink tracking-[0.15em]" style={{ fontWeight: 700 }}>{title.toUpperCase()}</h3>
          <p className="italic text-xs ocean mb-3">{subtitle}</p>
          <div className="flex items-baseline justify-between mb-2">
            <span className="font-display text-2xl ink" style={{ fontWeight: 700 }}>
              {visits} <span className="text-sm copper">/</span> <span className="copper">{target}</span>
            </span>
            <span className="font-display text-xs tracking-[0.2em] copper" style={{ fontWeight: 700 }}>{pct}%</span>
          </div>
          <div className="h-1.5  relative overflow-hidden">
            <div className="h-full bg-copper transition-all duration-700" style={{ width: `${pct}%` }}></div>
          </div>
        </div>
      </div>
    </div>
  );
}

function Leaderboard({ clients, visits, targets }) {
  const m = monthISO();
  const repStats = SALES_REPS.map(rep => {
    const repVisits = visits.filter(v => v.salesRep === rep && v.date && v.date.startsWith(m));
    const monthSales = repVisits.reduce((s, v) => s + (v.saleAmount || 0), 0);
    const target = targets[rep]?.revenue || 0;
    const pct = target > 0 ? Math.min(100, Math.round((monthSales / target) * 100)) : 0;
    const visitCount = repVisits.length;
    const visitTarget = targets[rep]?.visits || 0;
    const visitPct = visitTarget > 0 ? Math.min(100, Math.round((visitCount / visitTarget) * 100)) : 0;
    const sold = repVisits.filter(v => v.outcome === 'Sold In').length;
    return { rep, monthSales, target, pct, visitCount, visitTarget, visitPct, sold };
  }).sort((a, b) => b.monthSales - a.monthSales);

  return (
    <div className="space-y-2">
      {repStats.map((s, i) => (
        <div key={s.rep} className="border border p-3 hover:border-copper/40 transition-colors">
          <div className="flex items-center gap-3 mb-2">
            <div className="w-7 h-7 flex-shrink-0 flex items-center justify-center font-display text-xs bg-ink" style={{ color: '#FFFEF2', fontWeight: 700 }}>{i + 1}</div>
            <div className="flex-1 min-w-0">
              <div className="flex items-baseline justify-between gap-2">
                <h3 className="font-display text-sm tracking-[0.1em] ink truncate" style={{ fontWeight: 700 }}>{s.rep.toUpperCase()}</h3>
                <span className="font-display text-base copper flex-shrink-0" style={{ fontWeight: 700 }}>{ZAR(s.monthSales)}</span>
              </div>
              <p className="text-[10px] italic ocean truncate">{s.visitCount} visits · {s.sold} sold · target {ZAR(s.target)}</p>
            </div>
          </div>
          <div className="grid grid-cols-2 gap-2">
            <div>
              <div className="flex justify-between text-[9px] font-display tracking-[0.15em] mb-1">
                <span className="ocean" style={{ fontWeight: 600 }}>REVENUE</span>
                <span className="copper" style={{ fontWeight: 700 }}>{s.pct}%</span>
              </div>
              <div className="h-1 "><div className="h-full bg-copper" style={{ width: `${s.pct}%` }}></div></div>
            </div>
            <div>
              <div className="flex justify-between text-[9px] font-display tracking-[0.15em] mb-1">
                <span className="ocean" style={{ fontWeight: 600 }}>VISITS</span>
                <span className="copper" style={{ fontWeight: 700 }}>{s.visitPct}%</span>
              </div>
              <div className="h-1 "><div className="h-full" style={{ width: `${s.visitPct}%`, background: '#006C90' }}></div></div>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

function RecentVisits({ visits, clients }) {
  if (visits.length === 0) {
    return <p className="text-xs italic ocean py-6 text-center">No visits logged yet. Click "Log Visit" to begin.</p>;
  }
  return (
    <div className="space-y-2">
      {visits.map(v => {
        const c = clients.find(c => c.id === v.clientId);
        const outcomeColor = v.outcome === 'Sold In' ? '#006C90' : v.outcome === 'Rejected' ? '#9c2c2c' : '#D78433';
        return (
          <div key={v.id} className="border-l-2 pl-3 py-2" style={{ borderColor: outcomeColor }}>
            <div className="flex items-baseline justify-between">
              <span className="font-display text-xs ink tracking-wide" style={{ fontWeight: 700 }}>{c?.venue || 'Unknown'}</span>
              <span className="text-[10px] ocean">{v.date}</span>
            </div>
            <p className="text-[11px] italic" style={{ color: outcomeColor }}>{v.salesRep} · {v.outcome} {v.saleAmount > 0 && `· ${ZAR(v.saleAmount)}`}</p>
          </div>
        );
      })}
    </div>
  );
}

function PipelineGrid({ clients }) {
  const stages = ['New', 'Contacted', 'Qualified', 'Prospect', 'Converted', 'Lost'];
  const counts = stages.map(s => ({ stage: s, count: clients.filter(c => c.status === s).length }));
  const max = Math.max(...counts.map(c => c.count), 1);
  const colors = { New: '#006C90', Contacted: '#FDB940', Qualified: '#D78433', Prospect: '#003553', Converted: '#2d8659', Lost: '#9c2c2c' };
  return (
    <div className="grid grid-cols-3 md:grid-cols-6 gap-3">
      {counts.map(c => (
        <div key={c.stage} className="text-center p-2">
          <div className="font-display text-2xl md:text-3xl mb-1" style={{ fontWeight: 700, color: colors[c.stage] }}>{c.count}</div>
          <div className="font-display text-[9px] md:text-[10px] tracking-[0.15em] md:tracking-[0.25em] ocean mb-2" style={{ fontWeight: 600 }}>{c.stage.toUpperCase()}</div>
          <div className="h-1 ">
            <div className="h-full transition-all duration-700" style={{ width: `${(c.count / max) * 100}%`, background: colors[c.stage] }}></div>
          </div>
        </div>
      ))}
    </div>
  );
}

// =================== MANAGER PORTAL ===================
function ManagerPortal({ targets, saveTargets, clients, visits, askConfirm }) {
  const [draft, setDraft] = useState(targets);
  const [savedFlash, setSavedFlash] = useState(false);

  // Sync if targets change externally
  useEffect(() => { setDraft(targets); }, [targets]);

  const dirty = JSON.stringify(draft) !== JSON.stringify(targets);

  const setField = (rep, field, value) => {
    const v = Math.max(0, Number(value) || 0);
    setDraft(d => ({ ...d, [rep]: { ...d[rep], [field]: v } }));
  };

  const handleSave = async () => {
    await saveTargets(draft);
    setSavedFlash(true);
    setTimeout(() => setSavedFlash(false), 2000);
  };

  const handleReset = () => {
    askConfirm({
      title: 'Reset all targets?',
      message: 'This restores every rep to the default revenue, visits and channel targets.',
      confirmLabel: 'RESET TARGETS',
      danger: false,
      onConfirm: () => setDraft(DEFAULT_TARGETS),
    });
  };

  const handleRevert = () => setDraft(targets);

  // Live performance for context
  const m = monthISO();
  const repPerf = useMemo(() => {
    const out = {};
    SALES_REPS.forEach(rep => {
      const repVisits = visits.filter(v => v.salesRep === rep && v.date && v.date.startsWith(m));
      out[rep] = {
        revenue: repVisits.reduce((s, v) => s + (v.saleAmount || 0), 0),
        visits: repVisits.length,
        privateSales: repVisits.filter(v => v.channel === 'Private Sales').length,
        tradeRetail: repVisits.filter(v => v.channel === 'Trade Retail').length,
      };
    });
    return out;
  }, [visits, m]);

  // Team totals (from draft)
  const teamTotals = useMemo(() => SALES_REPS.reduce((acc, r) => ({
    revenue: acc.revenue + (draft[r]?.revenue || 0),
    visits: acc.visits + (draft[r]?.visits || 0),
    privateSales: acc.privateSales + (draft[r]?.privateSales || 0),
    tradeRetail: acc.tradeRetail + (draft[r]?.tradeRetail || 0),
  }), { revenue: 0, visits: 0, privateSales: 0, tradeRetail: 0 }), [draft]);

  // Client distribution per rep per channel
  const repBook = useMemo(() => {
    const out = {};
    SALES_REPS.forEach(rep => {
      const own = clients.filter(c => c.accountManager === rep);
      out[rep] = {
        total: own.length,
        privateSales: own.filter(c => c.channel === 'Private Sales').length,
        tradeRetail: own.filter(c => c.channel === 'Trade Retail').length,
      };
    });
    return out;
  }, [clients]);

  // Excel export
  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Clients
    const clientsRows = clients.map(c => ({
      'ID': c.id,
      'Venue': c.venue,
      'Channel': c.channel,
      'First Name': c.firstName,
      'Last Name': c.lastName,
      'Email': c.email,
      'Phone': c.phone,
      'Location': c.location,
      'Distributor': c.distributor,
      'Account Manager': c.accountManager,
      'Lead Source': c.leadSource,
      'Priority': c.priority,
      'Status': c.status,
      'Last Contacted': c.lastContacted,
      'Total Sales TD (R)': Number(c.totalSales || 0).toFixed(2),
      'Notes': c.notes,
    }));
    const wsClients = XLSX.utils.json_to_sheet(clientsRows);
    wsClients['!cols'] = [
      { wch: 5 }, { wch: 28 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
      { wch: 28 }, { wch: 16 }, { wch: 18 }, { wch: 14 }, { wch: 14 },
      { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 14 }, { wch: 16 }, { wch: 40 },
    ];
    XLSX.utils.book_append_sheet(wb, wsClients, 'Clients');

    // Sheet 2: Visits (header rows)
    const visitRows = visits.map(v => {
      const c = clients.find(cl => cl.id === v.clientId) || {};
      const itemSummary = (v.items && v.items.length > 0)
        ? v.items.map(it => `${it.qty}x ${it.name}`).join(', ')
        : '';
      return {
        'Visit ID': v.id,
        'Date': v.date,
        'Sales Rep': v.salesRep,
        'Venue': c.venue || 'Unknown',
        'Channel': v.channel || c.channel || '',
        'Location': c.location || '',
        'Outcome': v.outcome,
        'Order Total (R)': Number(v.saleAmount || 0).toFixed(2),
        'Items Summary': itemSummary,
        'Visit Notes': v.notes || '',
        'Follow-up Notes': v.followUp || '',
      };
    });
    const wsVisits = XLSX.utils.json_to_sheet(visitRows);
    wsVisits['!cols'] = [
      { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 28 }, { wch: 14 },
      { wch: 18 }, { wch: 18 }, { wch: 14 }, { wch: 50 }, { wch: 40 }, { wch: 40 },
    ];
    XLSX.utils.book_append_sheet(wb, wsVisits, 'Visits');

    // Sheet 3: Visit Line Items (one row per SKU sold)
    const lineRows = [];
    visits.forEach(v => {
      const c = clients.find(cl => cl.id === v.clientId) || {};
      if (v.items && v.items.length > 0) {
        v.items.forEach(it => {
          lineRows.push({
            'Visit ID': v.id,
            'Date': v.date,
            'Sales Rep': v.salesRep,
            'Venue': c.venue || 'Unknown',
            'Channel': v.channel || c.channel || '',
            'SKU': it.name,
            'Quantity': Number(it.qty) || 0,
            'Unit Price (R)': Number(it.unitPrice || 0).toFixed(2),
            'List Price (R)': Number(it.listPrice || 0).toFixed(2),
            'Discount (R)': (Number(it.listPrice || 0) - Number(it.unitPrice || 0)).toFixed(2),
            'Line Total (R)': Number(it.lineTotal || (Number(it.qty || 0) * Number(it.unitPrice || 0))).toFixed(2),
          });
        });
      }
    });
    const wsLines = XLSX.utils.json_to_sheet(lineRows.length > 0 ? lineRows : [{ 'Note': 'No line item data yet — start logging visits with SKUs.' }]);
    wsLines['!cols'] = [
      { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 28 }, { wch: 14 },
      { wch: 24 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 14 },
    ];
    XLSX.utils.book_append_sheet(wb, wsLines, 'SKU Line Items');

    // Sheet 4: Targets
    const targetRows = SALES_REPS.map(rep => ({
      'Sales Rep': rep,
      'Revenue Target (R)': Number(targets[rep]?.revenue || 0).toFixed(2),
      'Visits Target': targets[rep]?.visits || 0,
      'Private Sales Visits Target': targets[rep]?.privateSales || 0,
      'Trade Retail Visits Target': targets[rep]?.tradeRetail || 0,
    }));
    targetRows.push({
      'Sales Rep': 'TEAM TOTAL',
      'Revenue Target (R)': SALES_REPS.reduce((s, r) => s + (targets[r]?.revenue || 0), 0).toFixed(2),
      'Visits Target': SALES_REPS.reduce((s, r) => s + (targets[r]?.visits || 0), 0),
      'Private Sales Visits Target': SALES_REPS.reduce((s, r) => s + (targets[r]?.privateSales || 0), 0),
      'Trade Retail Visits Target': SALES_REPS.reduce((s, r) => s + (targets[r]?.tradeRetail || 0), 0),
    });
    const wsTargets = XLSX.utils.json_to_sheet(targetRows);
    wsTargets['!cols'] = [{ wch: 16 }, { wch: 20 }, { wch: 16 }, { wch: 28 }, { wch: 28 }];
    XLSX.utils.book_append_sheet(wb, wsTargets, 'Targets');

    // Sheet 5: Performance vs Target (current month)
    const monthLabel = new Date().toLocaleDateString('en-ZA', { month: 'long', year: 'numeric' });
    const m = monthISO();
    const perfRows = SALES_REPS.map(rep => {
      const repVisits = visits.filter(v => v.salesRep === rep && v.date && v.date.startsWith(m));
      const revenue = repVisits.reduce((s, v) => s + (v.saleAmount || 0), 0);
      const visitCount = repVisits.length;
      const privateCount = repVisits.filter(v => v.channel === 'Private Sales').length;
      const retailCount = repVisits.filter(v => v.channel === 'Trade Retail').length;
      const sold = repVisits.filter(v => v.outcome === 'Sold In').length;
      const t = targets[rep] || {};
      const pctOf = (a, b) => b > 0 ? ((a / b) * 100).toFixed(1) + '%' : '—';
      return {
        'Sales Rep': rep,
        'Month': monthLabel,
        'Revenue (R)': revenue.toFixed(2),
        'Revenue Target (R)': Number(t.revenue || 0).toFixed(2),
        'Revenue %': pctOf(revenue, t.revenue),
        'Visits': visitCount,
        'Visits Target': t.visits || 0,
        'Visits %': pctOf(visitCount, t.visits),
        'Private Sales Visits': privateCount,
        'Private Sales Target': t.privateSales || 0,
        'Trade Retail Visits': retailCount,
        'Trade Retail Target': t.tradeRetail || 0,
        'Sold In (count)': sold,
        'Conversion %': visitCount > 0 ? ((sold / visitCount) * 100).toFixed(1) + '%' : '—',
      };
    });
    const wsPerf = XLSX.utils.json_to_sheet(perfRows);
    wsPerf['!cols'] = Array(14).fill({ wch: 18 });
    XLSX.utils.book_append_sheet(wb, wsPerf, 'Performance');

    // SKU Catalogue reference
    const skuRows = SKU_CATALOGUE.map(s => ({
      'SKU': s.name,
      'List Price Trade Ex VAT (R)': s.price.toFixed(2),
    }));
    const wsSku = XLSX.utils.json_to_sheet(skuRows);
    wsSku['!cols'] = [{ wch: 28 }, { wch: 24 }];
    XLSX.utils.book_append_sheet(wb, wsSku, 'SKU Catalogue');

    const filename = `Avante_CRM_Export_${todayISO()}.xlsx`;
    XLSX.writeFile(wb, filename);
  };

  return (
    <div className="space-y-5 md:space-y-8 fade-up">
      {/* Hero — compact on mobile */}
      <div className="relative overflow-hidden border-2 border-ink" style={{ background: '#003553' }}>
        <RaysBackdrop opacity={0.1} />
        <div className="relative p-4 md:p-10">
          <div className="flex items-center gap-3 mb-3">
            <AvanteLogo height={44} />
            <div>
              <p className="font-display text-[9px] tracking-[0.3em]" style={{ color: '#FDB940', fontWeight: 600 }}>STRATEGIC EXECUTION</p>
              <h1 className="font-display text-xl md:text-4xl" style={{ color: '#FFFEF2', fontWeight: 700, letterSpacing: '0.06em' }}>MANAGER PORTAL</h1>
            </div>
          </div>
          {/* Action buttons — row on desktop, 2-col grid on mobile */}
          <div className="grid grid-cols-2 md:flex md:flex-row gap-2 mt-2">
            {savedFlash && (
              <span className="col-span-2 inline-flex items-center justify-center gap-2 px-3 py-2 font-display text-[10px] tracking-[0.2em]" style={{ color: '#003553', background: '#FDB940', fontWeight: 700 }}>
                <Save className="w-3.5 h-3.5" /> SAVED
              </span>
            )}
            <button onClick={handleReset} className="flex items-center justify-center gap-1.5 px-3 py-2.5 font-display text-[10px] tracking-[0.15em] border border-white/30" style={{ color: '#FFFEF2', fontWeight: 600 }}>
              <RotateCcw className="w-3.5 h-3.5" /> RESET
            </button>
            <button onClick={handleRevert} disabled={!dirty} className={`px-3 py-2.5 font-display text-[10px] tracking-[0.15em] border transition-colors ${dirty ? 'border-white/30' : 'border-white/10 opacity-30 cursor-not-allowed'}`} style={{ color: '#FFFEF2', fontWeight: 600 }}>
              REVERT
            </button>
            <button onClick={handleSave} disabled={!dirty} className={`col-span-2 md:col-span-1 flex items-center justify-center gap-2 px-5 py-2.5 font-display text-xs tracking-[0.2em] transition-colors ${dirty ? 'bg-copper' : 'bg-copper/40 cursor-not-allowed'}`} style={{ color: '#FFFEF2', fontWeight: 700 }}>
              <Save className="w-4 h-4" /> SAVE TARGETS
            </button>
          </div>
        </div>
      </div>

      {/* Data Export */}
      <div className="premium-card p-4 md:p-6 relative overflow-hidden">
        <div className="flex items-start gap-3 mb-3">
          <div className="w-10 h-10 flex-shrink-0 flex items-center justify-center bg-ink">
            <FileSpreadsheet className="w-5 h-5" style={{ color: '#FDB940' }} />
          </div>
          <div className="flex-1 min-w-0">
            <p className="font-display text-[9px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>DATA EXPORT</p>
            <h2 className="font-display text-lg ink" style={{ fontWeight: 700 }}>EXPORT TO EXCEL</h2>
            <p className="italic text-xs ocean mt-0.5">All clients, visits, SKUs, targets &amp; performance in one workbook.</p>
          </div>
        </div>
        <button onClick={exportToExcel} className="w-full flex items-center justify-center gap-2 bg-ink px-4 py-3 font-display text-xs tracking-[0.25em]" style={{ color: '#FFFEF2', fontWeight: 700 }}>
          <Download className="w-4 h-4" /> DOWNLOAD .XLSX
        </button>
        <div className="mt-3 pt-3 border-t border grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
          <div>
            <p className="font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>CLIENTS IN DB</p>
            <p className="font-display text-lg ink" style={{ fontWeight: 700 }}>{clients.length}</p>
          </div>
          <div>
            <p className="font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>TOTAL VISITS</p>
            <p className="font-display text-lg ink" style={{ fontWeight: 700 }}>{visits.length}</p>
          </div>
          <div>
            <p className="font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>SKU LINES SOLD</p>
            <p className="font-display text-lg ink" style={{ fontWeight: 700 }}>
              {visits.reduce((s, v) => s + (v.items?.length || 0), 0)}
            </p>
          </div>
          <div>
            <p className="font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>ALL-TIME REVENUE</p>
            <p className="font-display text-lg copper" style={{ fontWeight: 700 }}>
              {ZAR(visits.reduce((s, v) => s + (v.saleAmount || 0), 0))}
            </p>
          </div>
        </div>
      </div>

      {/* Team aggregate cards */}
      <div>
        <div className="flex items-center gap-3 mb-3">
          <div className="w-2 h-2 bg-copper diamond-clip"></div>
          <h2 className="font-display text-sm tracking-[0.3em] ink" style={{ fontWeight: 700 }}>TEAM AGGREGATE TARGETS</h2>
        </div>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <AggregateCard label="Total Revenue Target" value={ZAR(teamTotals.revenue)} icon={DollarSign} color="#D78433" />
          <AggregateCard label="Total Visits Target" value={teamTotals.visits} icon={Activity} color="#006C90" />
          <AggregateCard label="Private Sales Visits" value={teamTotals.privateSales} icon={Wine} color="#FDB940" />
          <AggregateCard label="Trade Retail Visits" value={teamTotals.tradeRetail} icon={Briefcase} color="#003553" />
        </div>
      </div>

      {/* Per-rep target editor cards */}
      <div>
        <div className="flex items-center gap-3 mb-4">
          <div className="w-2 h-2 bg-copper diamond-clip"></div>
          <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>INDIVIDUAL REP TARGETS &amp; KPI</h2>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {SALES_REPS.map(rep => (
            <RepTargetCard
              key={rep}
              rep={rep}
              draft={draft[rep]}
              perf={repPerf[rep]}
              book={repBook[rep]}
              onChange={(field, val) => setField(rep, field, val)}
            />
          ))}
        </div>
      </div>

      {/* Quick reference: client distribution — scrollable on mobile */}
      <div className="premium-card p-4 md:p-6">
        <div className="flex items-center gap-3 mb-3">
          <div className="w-2 h-2 bg-copper diamond-clip"></div>
          <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>CLIENT BOOK DISTRIBUTION</h2>
        </div>
        {/* Mobile: cards */}
        <div className="md:hidden space-y-3">
          {SALES_REPS.map(rep => {
            const b = repBook[rep];
            const targetVisits = draft[rep]?.visits || 0;
            const ratio = b.total > 0 ? (targetVisits / b.total).toFixed(2) : '—';
            return (
              <div key={rep} className="border border p-3">
                <div className="flex items-center justify-between mb-2">
                  <span className="font-display ink text-sm tracking-wider" style={{ fontWeight: 700 }}>{rep.toUpperCase()}</span>
                  <span className="font-display text-xs copper" style={{ fontWeight: 700 }}>{ratio}× visits/client</span>
                </div>
                <div className="grid grid-cols-3 gap-2 text-xs">
                  <div><p className="font-display text-[9px] ocean tracking-wider" style={{ fontWeight: 600 }}>TOTAL</p><p className="ink font-display" style={{ fontWeight: 700 }}>{b.total}</p></div>
                  <div><p className="font-display text-[9px] ocean tracking-wider" style={{ fontWeight: 600 }}>PRIVATE</p><p className="ink font-display" style={{ fontWeight: 700 }}>{b.privateSales}</p></div>
                  <div><p className="font-display text-[9px] ocean tracking-wider" style={{ fontWeight: 600 }}>RETAIL</p><p className="ink font-display" style={{ fontWeight: 700 }}>{b.tradeRetail}</p></div>
                </div>
              </div>
            );
          })}
        </div>
        {/* Desktop: table */}
        <div className="hidden md:block overflow-x-auto scrollbar-thin">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b-2 border-ink">
                {['Sales Rep', 'Total', 'Private Sales', 'Trade Retail', 'Visit/Client Ratio'].map(h => (
                  <th key={h} className="px-3 py-2 text-left font-display text-[10px] tracking-[0.2em] copper" style={{ fontWeight: 600 }}>{h.toUpperCase()}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {SALES_REPS.map(rep => {
                const b = repBook[rep];
                const targetVisits = draft[rep]?.visits || 0;
                const ratio = b.total > 0 ? (targetVisits / b.total).toFixed(2) : '—';
                return (
                  <tr key={rep} className="border-b border">
                    <td className="px-3 py-3 font-display ink tracking-wider" style={{ fontWeight: 700 }}>{rep.toUpperCase()}</td>
                    <td className="px-3 py-3 ink">{b.total}</td>
                    <td className="px-3 py-3 ink">{b.privateSales} <span className="ocean italic text-xs">venues</span></td>
                    <td className="px-3 py-3 ink">{b.tradeRetail} <span className="ocean italic text-xs">venues</span></td>
                    <td className="px-3 py-3 copper" style={{ fontWeight: 700 }}>{ratio}<span className="text-xs ocean italic"> per client / mo</span></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Save reminder bar — sits above mobile bottom nav */}
      {dirty && (
        <div className="sticky bottom-20 md:bottom-4 z-30 max-w-2xl mx-auto">
          <div className="flex items-center justify-between gap-3 p-3 md:p-4 border-2 border-copper shadow-2xl" style={{ background: '#003553' }}>
            <div className="flex items-center gap-2">
              <div className="w-2 h-2 bg-gold diamond-clip animate-pulse flex-shrink-0"></div>
              <p className="font-display text-[9px] md:text-xs tracking-[0.15em]" style={{ color: '#FFFEF2', fontWeight: 600 }}>UNSAVED CHANGES</p>
            </div>
            <div className="flex items-center gap-2">
              <button onClick={handleRevert} className="px-2 md:px-3 py-2 font-display text-[9px] tracking-[0.15em]  " style={{ fontWeight: 600 }}>REVERT</button>
              <button onClick={handleSave} className="flex items-center gap-1.5 bg-copper px-3 md:px-4 py-2 font-display text-[9px] md:text-[10px] tracking-[0.2em]" style={{ color: '#FFFEF2', fontWeight: 700 }}>
                <Save className="w-3.5 h-3.5" /> SAVE
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function AggregateCard({ label, value, icon: Icon, color }) {
  return (
    <div className="premium-card p-5 relative overflow-hidden">
      <div className="absolute top-0 right-0 w-16 h-16 opacity-[0.06]" style={{ background: color, clipPath: 'polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)', transform: 'translate(30%, -30%)' }}></div>
      <div className="flex items-center justify-between mb-3 relative">
        <span className="font-display text-[10px] tracking-[0.25em]" style={{ color, fontWeight: 600 }}>{label.toUpperCase()}</span>
        <Icon className="w-4 h-4" style={{ color }} />
      </div>
      <div className="font-display text-2xl ink" style={{ fontWeight: 700 }}>{value}</div>
    </div>
  );
}

function RepTargetCard({ rep, draft, perf, book, onChange }) {
  const fields = [
    { key: 'revenue', label: 'Revenue Target', icon: DollarSign, prefix: 'R', perfKey: 'revenue', isCurrency: true, hint: 'Monthly Rand target' },
    { key: 'visits', label: 'Total Visits', icon: Activity, perfKey: 'visits', hint: 'Total venue visits' },
    { key: 'privateSales', label: 'Private Sales', icon: Wine, perfKey: 'privateSales', hint: 'On-con visits KPI' },
    { key: 'tradeRetail', label: 'Trade Retail', icon: Briefcase, perfKey: 'tradeRetail', hint: 'Off-con visits KPI' },
  ];

  return (
    <div className="premium-card p-4 md:p-6 relative">
      <div className="flex items-center justify-between pb-3 mb-3 border-b border">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 flex items-center justify-center bg-ink">
            <span className="font-display text-sm" style={{ color: '#FDB940', fontWeight: 700 }}>{rep[0]}</span>
          </div>
          <div>
            <h3 className="font-display text-base tracking-[0.1em] ink" style={{ fontWeight: 700 }}>{rep.toUpperCase()}</h3>
            <p className="text-[10px] italic ocean">{book.total} clients</p>
          </div>
        </div>
        <Target className="w-4 h-4 copper" />
      </div>

      <div className="space-y-3">
        {fields.map(f => {
          const Icon = f.icon;
          const val = draft?.[f.key] || 0;
          const actual = perf?.[f.perfKey] || 0;
          const pct = val > 0 ? Math.min(100, Math.round((actual / val) * 100)) : 0;
          return (
            <div key={f.key}>
              <div className="flex items-center gap-1.5 mb-1">
                <Icon className="w-3 h-3 copper" />
                <label className="font-display text-[9px] tracking-[0.2em] copper" style={{ fontWeight: 600 }}>{f.label.toUpperCase()}</label>
              </div>
              <div className="flex items-center gap-2">
                {f.prefix && <span className="font-display ink text-sm flex-shrink-0" style={{ fontWeight: 700 }}>{f.prefix}</span>}
                <input
                  type="number"
                  min="0"
                  value={val}
                  onChange={(e) => onChange(f.key, e.target.value)}
                  className="flex-1 px-2 py-1.5 border border bg-cream font-display text-sm focus:outline-none focus:border-copper ink"
                  style={{ fontWeight: 700 }}
                />
              </div>
              <div className="flex items-center justify-between mt-1.5 text-[10px]">
                <span className="ocean italic">{f.hint} · actual: <span className="ink font-display" style={{ fontWeight: 700 }}>{f.isCurrency ? ZAR(actual) : actual}</span></span>
                <span className="font-display tracking-wider copper flex-shrink-0" style={{ fontWeight: 700 }}>{pct}%</span>
              </div>
              <div className="h-1  mt-1">
                <div className="h-full bg-copper transition-all duration-500" style={{ width: `${pct}%` }}></div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// =================== Leads Page ===================
function LeadsPage({ clients, visits, updateClient, onSelect }) {
  const [search, setSearch] = useState('');
  const [filterRep, setFilterRep] = useState('All');
  const [filterChannel, setFilterChannel] = useState('All');

  // Alphabetically sorted + filtered list
  const filtered = useMemo(() => {
    return clients
      .filter(c => {
        if (filterRep !== 'All' && c.accountManager !== filterRep) return false;
        if (filterChannel !== 'All' && c.channel !== filterChannel) return false;
        if (search.trim()) {
          const q = search.trim().toLowerCase();
          const hay = [c.venue, c.firstName, c.lastName, c.location, c.email, c.phone, c.notes]
            .filter(Boolean).join(' ').toLowerCase();
          if (!hay.includes(q)) return false;
        }
        return true;
      })
      .sort((a, b) => (a.venue || '').localeCompare(b.venue || ''));
  }, [clients, filterRep, filterChannel, search]);

  return (
    <div className="space-y-4 fade-up">
      {/* Page header */}
      <div className="pb-3 border-b border">
        <p className="font-display text-[9px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>CLIENT DATABASE</p>
        <h1 className="font-display mt-1 ink" style={{ fontWeight: 700, fontSize: 28 }}>LEADS &amp; CLIENTS</h1>
        <p className="italic ocean" style={{ fontSize: 12, marginTop: 2 }}>{clients.length} venues tracked</p>
      </div>

      {/* Search + filters */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {/* Search bar — plain input, no absolute positioning */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, border: '1px solid rgba(0,53,83,0.25)', background: '#FFFEF2', padding: '0 12px' }}>
          <Search style={{ width: 16, height: 16, color: '#006C90', flexShrink: 0 }} />
          <input
            type="text"
            placeholder="Search venue, contact, location..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={{ flex: 1, border: 'none', background: 'transparent', padding: '12px 0', fontFamily: "'Libre Baskerville', Georgia, serif", fontSize: 14, color: '#003553', outline: 'none' }}
          />
          {search && (
            <button onClick={() => setSearch('')} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#006C90', fontSize: 18, lineHeight: 1 }}>×</button>
          )}
        </div>
        {/* Rep + Channel filters */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
          <FilterSelect label="Rep" value={filterRep} onChange={setFilterRep} options={['All', ...SALES_REPS, 'Unassigned']} />
          <FilterSelect label="Channel" value={filterChannel} onChange={setFilterChannel} options={['All', ...CHANNELS]} />
        </div>
        <p style={{ fontSize: 11, fontStyle: 'italic', color: '#006C90', margin: 0 }}>
          {filtered.length} of {clients.length} clients{search ? ` matching "${search}"` : ''}
        </p>
      </div>

      {/* Client list — single unified view, works on all screen sizes */}
      <div className="premium-card" style={{ overflow: 'hidden' }}>
        {filtered.length === 0 ? (
          <div style={{ padding: 40, textAlign: 'center' }}>
            <p style={{ fontStyle: 'italic', color: '#006C90', fontSize: 14 }}>
              {search ? `No clients found matching "${search}"` : 'No clients match your filters.'}
            </p>
            {search && (
              <button onClick={() => setSearch('')}
                style={{ marginTop: 12, padding: '8px 20px', background: '#D78433', color: '#FFFEF2', border: 'none', cursor: 'pointer', fontFamily: "'Cinzel', serif", fontSize: 10, letterSpacing: '0.2em', fontWeight: 700 }}>
                CLEAR SEARCH
              </button>
            )}
          </div>
        ) : (
          <div style={{ maxHeight: 640, overflowY: 'auto' }}>
            {filtered.map((c, i) => (
              <div key={c.id} onClick={() => onSelect(c)}
                style={{
                  display: 'flex', alignItems: 'center', gap: 12, padding: '12px 16px',
                  borderTop: i === 0 ? 'none' : '1px solid rgba(0,53,83,0.1)',
                  borderLeft: `3px solid ${c.channel === 'Private Sales' ? '#D78433' : '#006C90'}`,
                  cursor: 'pointer', transition: 'background 0.12s',
                  background: '#FFFEF2',
                }}
                onMouseEnter={e => e.currentTarget.style.background = 'rgba(0,53,83,0.04)'}
                onMouseLeave={e => e.currentTarget.style.background = '#FFFEF2'}
              >
                {/* Venue + notes */}
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                    <span className="font-display ink" style={{ fontWeight: 700, fontSize: 13 }}>{c.venue}</span>
                    <StatusBadge status={c.status} />
                  </div>
                  <p style={{ fontSize: 11, color: '#006C90', fontStyle: 'italic', margin: '2px 0 0', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {[c.firstName, c.lastName].filter(Boolean).join(' ')}
                    {c.location ? ((c.firstName || c.lastName) ? ' · ' : '') + c.location : ''}
                  </p>
                </div>
                {/* Channel badge */}
                <span style={{ fontSize: 9, fontFamily: "'Cinzel', serif", fontWeight: 600, letterSpacing: '0.15em', padding: '2px 6px', border: `1px solid ${c.channel === 'Private Sales' ? '#D78433' : '#006C90'}`, color: c.channel === 'Private Sales' ? '#D78433' : '#006C90', flexShrink: 0 }}>
                  {c.channel === 'Private Sales' ? 'PRIVATE' : c.channel === 'Trade Retail' ? 'RETAIL' : '—'}
                </span>
                {/* Rep */}
                <span className="font-display" style={{ fontSize: 11, fontWeight: 700, color: '#003553', flexShrink: 0, minWidth: 48, textAlign: 'right' }}>
                  {c.accountManager === 'Unassigned' ? '—' : c.accountManager}
                </span>
                {/* Sales */}
                {c.totalSales > 0 && (
                  <span className="font-display copper" style={{ fontSize: 11, fontWeight: 700, flexShrink: 0 }}>{ZAR(c.totalSales)}</span>
                )}
                <ChevronRight style={{ width: 16, height: 16, color: '#D78433', flexShrink: 0 }} />
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function FilterSelect({ label, value, onChange, options }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
      <label className="font-display ocean" style={{ fontSize: 9, fontWeight: 600, letterSpacing: '0.2em' }}>
        {label.toUpperCase()}
      </label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        style={{ width: '100%', padding: '8px 10px', border: '1px solid rgba(0,53,83,0.25)', background: '#FFFEF2', fontFamily: "'Libre Baskerville', Georgia, serif", fontSize: 13, color: '#003553', cursor: 'pointer', outline: 'none' }}
      >
        {options.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
  );
}

function StatusBadge({ status }) {
  const colors = {
    New: { bg: '#006C90', text: '#FFFEF2' },
    Contacted: { bg: '#FDB940', text: '#003553' },
    Qualified: { bg: '#D78433', text: '#FFFEF2' },
    Prospect: { bg: '#003553', text: '#FFFEF2' },
    Converted: { bg: '#2d8659', text: '#FFFEF2' },
    Lost: { bg: '#9c2c2c', text: '#FFFEF2' },
  };
  const c = colors[status] || colors.New;
  return (
    <span className="inline-block px-2 py-0.5 text-[10px] font-display tracking-wider" style={{ background: c.bg, color: c.text, fontWeight: 600 }}>
      {(status || 'NEW').toUpperCase()}
    </span>
  );
}

// =================== Visits Page ===================
function VisitsPage({ visits, clients, onLog, onEdit, onDelete }) {
  const [filterRep, setFilterRep] = useState('All');
  const filtered = visits.filter(v => filterRep === 'All' || v.salesRep === filterRep).sort((a, b) => (b.date || '').localeCompare(a.date || ''));

  return (
    <div className="space-y-6 fade-up">
      {/* Page header */}
      <div className="pb-3 border-b border">
        <p className="font-display text-[9px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>VISIT HISTORY</p>
        <div className="flex items-center justify-between gap-3 mt-1 flex-wrap">
          <h1 className="font-display text-2xl md:text-3xl ink" style={{ fontWeight: 700 }}>VISIT LOG</h1>
          <div className="flex items-center gap-3 flex-wrap">
            <RepToggle active={filterRep} onChange={setFilterRep} />
            <button onClick={onLog}
              className="flex items-center gap-2 px-4 py-2.5 font-display text-[10px] tracking-[0.2em]"
              style={{ background: '#D78433', color: '#FFFEF2', fontWeight: 700 }}>
              <Plus className="w-4 h-4" /> LOG VISIT
            </button>
          </div>
        </div>
      </div>

      {filtered.length === 0 ? (
        <div className="premium-card p-10 text-center">
          <ClipboardList className="w-10 h-10 mx-auto mb-3 ocean opacity-40" />
          <p className="font-display text-sm tracking-[0.2em] ink mb-2" style={{ fontWeight: 700 }}>NO VISITS LOGGED YET</p>
          <p className="italic text-sm ocean mb-5">Start tracking your call cycle.</p>
          <button onClick={onLog} className="bg-copper px-6 py-3 font-display text-xs tracking-[0.25em]" style={{ color: '#FFFEF2', fontWeight: 700 }}>LOG YOUR FIRST VISIT</button>
        </div>
      ) : (
        <div className="space-y-2">
          {filtered.map(v => {
            const c = clients.find(c => c.id === v.clientId);
            const outcomeColor = v.outcome === 'Sold In' ? '#2d8659' : v.outcome === 'Rejected' ? '#9c2c2c' : '#D78433';
            return (
              <div key={v.id} className="premium-card border-l-4" style={{ borderLeftColor: outcomeColor }}>
                {/* Main row — always visible */}
                <div className="p-3 md:p-5 flex items-start justify-between gap-3">
                  <div className="flex-1 min-w-0">
                    {/* Venue + outcome */}
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="font-display ink text-sm truncate" style={{ fontWeight: 700 }}>{c?.venue || 'Unknown'}</span>
                      <span className="text-[9px] font-display tracking-wider px-1.5 py-0.5 flex-shrink-0" style={{ background: outcomeColor, color: '#FFFEF2', fontWeight: 600 }}>{v.outcome.toUpperCase()}</span>
                    </div>
                    {/* Meta line */}
                    <p className="text-[11px] ocean italic mt-0.5">
                      {v.date} · {v.salesRep}{c?.location ? ' · ' + c.location : ''}
                    </p>
                    {/* Sale amount */}
                    {v.saleAmount > 0 && (
                      <p className="font-display copper text-base mt-1" style={{ fontWeight: 700 }}>{ZAR(v.saleAmount)}</p>
                    )}
                    {/* SKU summary on mobile */}
                    {v.items && v.items.length > 0 && (
                      <p className="text-[10px] ocean mt-1">{v.items.length} SKU{v.items.length > 1 ? 's' : ''}: {v.items.map(it => `${it.qty}× ${it.name}`).join(', ').slice(0, 60)}{v.items.map(it => `${it.qty}× ${it.name}`).join(', ').length > 60 ? '…' : ''}</p>
                    )}
                    {/* Notes preview */}
                    {v.notes && <p className="text-[10px] italic ink mt-1 truncate">"{v.notes}"</p>}
                  </div>
                  {/* Actions */}
                  <div className="flex items-center gap-1 flex-shrink-0">
                    <button type="button" onClick={() => onEdit(v)} className="p-2 hover: border border">
                      <Edit2 className="w-3.5 h-3.5 ocean" />
                    </button>
                    <button type="button" onClick={() => onDelete(v.id)} className="p-2  border border">
                      <Trash2 className="w-3.5 h-3.5" style={{ color: '#9c2c2c' }} />
                    </button>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// =================== Log Visit Modal ===================
function LogVisitModal({ clients, onClose, onSubmit, onRequestNewClient, existingVisit }) {
  const isEdit = !!existingVisit;
  const [salesRep, setSalesRep] = useState(existingVisit?.salesRep || 'Alex');
  const [clientId, setClientId] = useState(existingVisit?.clientId ? String(existingVisit.clientId) : '');
  const [date, setDate] = useState(existingVisit?.date || todayISO());
  const [outcome, setOutcome] = useState(existingVisit?.outcome || 'Met / Discussion');
  const [items, setItems] = useState(existingVisit?.items ? existingVisit.items.map(it => ({ ...it })) : []);
  const [notes, setNotes] = useState(existingVisit?.notes || '');
  const [followUp, setFollowUp] = useState(existingVisit?.followUp || '');
  const [search, setSearch] = useState('');
  const [skuPickerOpen, setSkuPickerOpen] = useState(false);

  // STRICT filter: only this rep's clients (no Unassigned leakage)
  // BUT in edit mode, also include the visit's existing client even if reassigned to someone else.
  const repClients = useMemo(() => {
    let list = clients.filter(c => c.accountManager === salesRep);
    // Make sure the currently-selected client is in the list, even if reassigned away from this rep
    if (isEdit && existingVisit?.clientId) {
      const existingClient = clients.find(c => c.id === existingVisit.clientId);
      if (existingClient && !list.some(c => c.id === existingClient.id)) {
        list = [existingClient, ...list];
      }
    }
    if (search) {
      const q = search.toLowerCase();
      list = list.filter(c => `${c.venue} ${c.location}`.toLowerCase().includes(q));
    }
    return list.sort((a, b) => (a.venue || '').localeCompare(b.venue || ''));
  }, [clients, salesRep, search, isEdit, existingVisit]);

  const selectedClient = clients.find(c => c.id === Number(clientId));

  // Compute totals
  const orderTotal = useMemo(
    () => items.reduce((s, it) => s + (Number(it.unitPrice) || 0) * (Number(it.qty) || 0), 0),
    [items]
  );

  // SKU operations
  const addSku = (sku) => {
    if (items.find(it => it.skuId === sku.id)) {
      // already in list — bump qty
      setItems(items.map(it => it.skuId === sku.id ? { ...it, qty: it.qty + 1 } : it));
    } else {
      setItems([...items, { skuId: sku.id, name: sku.name, qty: 1, unitPrice: sku.price, listPrice: sku.price }]);
    }
    setSkuPickerOpen(false);
  };
  const updateItem = (skuId, field, value) => {
    setItems(items.map(it => it.skuId === skuId ? { ...it, [field]: value } : it));
  };
  const removeItem = (skuId) => setItems(items.filter(it => it.skuId !== skuId));

  const [validationError, setValidationError] = useState('');

  // Build a mailto: link with the order details prefilled.
  // Opens the user's default email client (works on desktop + mobile, no SMTP needed).
  // Compose the order email body. Pure function — no side effects.
  // Used by both the preview and the actual mailto link.
  const composeOrderEmail = () => {
    const c = selectedClient;
    const venueLine = c?.venue || '';
    const contactName = [c?.firstName, c?.lastName].filter(Boolean).join(' ');
    const orderTotal = items.reduce((s, it) => s + (Number(it.unitPrice) || 0) * (Number(it.qty) || 0), 0);
    const senderEmail = REP_EMAILS[salesRep] || '';

    const subject = `Avante Cape Brandy — Order Confirmation${venueLine ? ' · ' + venueLine : ''}${date ? ' · ' + date : ''}`;

    const lines = [];
    lines.push(`Dear ${contactName || 'Sir / Madam'},`);
    lines.push('');
    lines.push(`Thank you for your order with Avante Cape Brandy. Please find the details below for your records:`);
    lines.push('');
    lines.push('─────────────────────────────────────────');
    lines.push('VISIT DETAILS');
    lines.push('─────────────────────────────────────────');
    lines.push(`Venue:          ${venueLine || '(not specified)'}`);
    if (c?.location) lines.push(`Location:       ${c.location}`);
    if (contactName) lines.push(`Contact:        ${contactName}`);
    lines.push(`Date:           ${date || '(not specified)'}`);
    lines.push(`Outcome:        ${outcome}`);
    lines.push(`Sales Rep:      ${salesRep} — Avante Cape Brandy`);
    lines.push('');

    if (items.length > 0) {
      lines.push('─────────────────────────────────────────');
      lines.push('ORDER PLACED');
      lines.push('─────────────────────────────────────────');
      lines.push('Qty   SKU                                    Unit (R)      Line (R)');
      lines.push('-----------------------------------------------------------------------');
      items.forEach(it => {
        const qty = String(Number(it.qty) || 0).padEnd(5);
        const name = (it.name || '').padEnd(38).slice(0, 38);
        const unit = ('R ' + (Number(it.unitPrice) || 0).toFixed(2)).padStart(11);
        const line = ('R ' + ((Number(it.unitPrice) || 0) * (Number(it.qty) || 0)).toFixed(2)).padStart(13);
        lines.push(`${qty} ${name} ${unit}  ${line}`);
        const discounted = Number(it.unitPrice) < Number(it.listPrice);
        if (discounted) {
          lines.push(`      (discounted from R ${Number(it.listPrice).toFixed(2)} — saved R ${(Number(it.listPrice) - Number(it.unitPrice)).toFixed(2)} per unit)`);
        }
      });
      lines.push('-----------------------------------------------------------------------');
      lines.push(`${''.padEnd(45)}TOTAL EX VAT  R ${orderTotal.toFixed(2)}`);
      lines.push('');
      lines.push('All prices shown are Trade Ex VAT in South African Rand.');
      lines.push('');
    }

    if (notes) {
      lines.push('─────────────────────────────────────────');
      lines.push('NOTES FROM THE VISIT');
      lines.push('─────────────────────────────────────────');
      lines.push(notes);
      lines.push('');
    }

    if (followUp) {
      lines.push('─────────────────────────────────────────');
      lines.push('FOLLOW-UP');
      lines.push('─────────────────────────────────────────');
      lines.push(followUp);
      lines.push('');
    }

    lines.push('Should you need any further information, please do not hesitate to contact me.');
    lines.push('');
    lines.push('Kind regards,');
    lines.push(salesRep);
    if (senderEmail) lines.push(senderEmail);
    lines.push('Avante Cape Brandy');
    lines.push('— Dare to Forward —');

    return { subject, body: lines.join('\n'), defaultRecipient: c?.email || '', orderTotal };
  };

  // State for the email-recipient modal
  const [emailModalOpen, setEmailModalOpen] = useState(false);

  const handleEmailOrder = () => {
    setValidationError('');
    if (!clientId) {
      setValidationError('Select a client first — the email needs an order to summarise.');
      return;
    }
    if (items.length === 0 && !notes && !followUp) {
      setValidationError('Nothing to email yet — add SKUs, notes, or follow-up first.');
      return;
    }
    setEmailModalOpen(true);
  };

  // Build the mailto URL. Recipients use comma separators per RFC 6068
  // and must NOT be URL-encoded (encoding the comma breaks mail clients).
  // Subject and body do get encoded, but with %20 spaces (not '+') for compatibility.
  const buildMailtoUrl = (recipient) => {
    const { subject, body } = composeOrderEmail();
    // Recipients: split, trim, validate, rejoin with commas — don't encodeURIComponent
    const recipientList = recipient.split(/[,;]/).map(s => s.trim()).filter(Boolean).join(',');
    const enc = (s) => encodeURIComponent(s).replace(/'/g, '%27');
    return `mailto:${recipientList}?subject=${enc(subject)}&body=${enc(body)}`;
  };

  // Fire the mailto link. Called when the user clicks "Open in Mail App".
  // We build a real <a> element and click it — this is the most reliable
  // cross-browser/cross-iframe approach. Sandboxed iframes block
  // window.location.href = mailto:... but anchor clicks still work.
  const sendOrderEmail = (recipient) => {
    const url = buildMailtoUrl(recipient);
    try {
      const a = document.createElement('a');
      a.href = url;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      // Some browsers need the anchor to be in the DOM briefly
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    } catch (err) {
      console.error('[LogVisit] mailto failed', err);
      setValidationError('Could not open email client automatically. Use the link to copy or open manually.');
    }
  };

  const handleSave = () => {
    console.log('[LogVisit] handleSave clicked', { clientId, salesRep, items, date, outcome });
    setValidationError('');
    if (!clientId) {
      setValidationError('Please select a client before saving.');
      return;
    }
    const payload = {
      salesRep,
      clientId: Number(clientId),
      channel: selectedClient?.channel || '',
      date, outcome,
      items: items.map(it => ({
        skuId: it.skuId,
        name: it.name,
        qty: Number(it.qty) || 0,
        unitPrice: Number(it.unitPrice) || 0,
        listPrice: Number(it.listPrice) || 0,
        lineTotal: (Number(it.unitPrice) || 0) * (Number(it.qty) || 0),
      })),
      notes, followUp,
    };
    console.log('[LogVisit] calling onSubmit with payload', payload);
    try {
      onSubmit(payload);
    } catch (err) {
      console.error('[LogVisit] onSubmit threw:', err);
      setValidationError('Save failed: ' + (err?.message || String(err)));
    }
  };

  const requestAddClient = () => {
    onRequestNewClient(salesRep, (created) => {
      setClientId(String(created.id));
    });
  };

  return (
    <div onClick={onClose} style={{ position:"fixed", inset:0, zIndex:50, display:"flex", alignItems:"center", justifyContent:"center", padding:16, background:"rgba(0,53,83,0.78)", backdropFilter:"blur(4px)", overflowY:"auto" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background:"#FFFEF2", width:"100%", maxWidth:768, maxHeight:"92vh", overflowY:"auto", border:"2px solid #D78433", position:"relative" }}>
        <div className="absolute top-0 left-0 right-0 h-1 bg-copper z-10"></div>
        {/* Header */}
        <div className="p-4 md:p-6 relative overflow-hidden" style={{ background: '#003553' }}>
          <RaysBackdrop opacity={0.06} />
          <div className="flex items-center justify-between gap-4 relative">
            <div className="flex items-center gap-3">
              <AvanteLogo height={36} />
              <div>
                <p className="font-display text-[9px] tracking-[0.3em]" style={{ color: '#FDB940', fontWeight: 600 }}>{isEdit ? 'AMEND RECORD' : 'DARE TO FORWARD'}</p>
                <h2 className="font-display text-xl md:text-2xl mt-0.5" style={{ color: '#FFFEF2', fontWeight: 700, letterSpacing: '0.08em' }}>{isEdit ? 'EDIT VISIT' : 'LOG A VISIT'}</h2>
              </div>
            </div>
            <button onClick={onClose} className="p-2 hover:opacity-70" style={{ color: '#FFFEF2' }}><X className="w-5 h-5" /></button>
          </div>
        </div>

        <div className="p-4 md:p-6 space-y-4">
          {/* Sales rep */}
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>SALES MEMBER</label>
            <div className="grid grid-cols-3 gap-2">
              {SALES_REPS.map(r => (
                <button key={r} type="button" onClick={() => { setSalesRep(r); setClientId(''); }} className={`py-3 font-display text-xs tracking-[0.2em] border transition-all ${salesRep === r ? 'bg-ink border-ink' : 'bg-cream border hover:border-copper'}`} style={{ color: salesRep === r ? '#FFFEF2' : '#003553', fontWeight: 700 }}>
                  {r.toUpperCase()}
                </button>
              ))}
            </div>
          </div>

          {/* Client picker */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>CLIENT ({repClients.length})</label>
              <button type="button" onClick={requestAddClient} className="flex items-center gap-1.5 text-[10px] font-display tracking-[0.2em] copper" style={{ fontWeight: 700 }}>
                <UserPlus className="w-3.5 h-3.5" /> ADD NEW
              </button>
            </div>
            <div className="relative mb-2">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 ocean" />
              <input type="text" placeholder="Filter venues..." value={search} onChange={(e) => setSearch(e.target.value)} className="w-full pl-10 pr-3 py-2 border border bg-cream text-sm focus:outline-none focus:border-copper" />
            </div>
            <select value={clientId} onChange={(e) => setClientId(e.target.value)} className="w-full px-3 py-3 border border bg-cream font-body text-sm focus:outline-none focus:border-copper">
              <option value="">— Select a venue —</option>
              {repClients.map(c => (
                <option key={c.id} value={c.id}>{c.venue}{c.location ? ' · ' + c.location : ''}</option>
              ))}
            </select>
            {selectedClient && (
              <div className="mt-2 p-3  border-l-2 border-copper text-xs">
                <p className="ink font-display" style={{ fontWeight: 700 }}>{selectedClient.venue}</p>
                <p className="ocean italic">{selectedClient.location} · {selectedClient.channel || 'No channel'} · {selectedClient.status}</p>
              </div>
            )}
          </div>

          {/* Date + Outcome — side by side on all screens */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>DATE</label>
              <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="w-full px-3 py-3 border border bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>OUTCOME</label>
              <select value={outcome} onChange={(e) => setOutcome(e.target.value)} className="w-full px-3 py-3 border border bg-cream font-body text-sm focus:outline-none focus:border-copper">
                <option>Met / Discussion</option>
                <option>Sold In</option>
                <option>Sample Drop</option>
                <option>Quoted</option>
                <option>Follow-Up Required</option>
                <option>Rejected</option>
                <option>No Show</option>
              </select>
            </div>
          </div>

          {/* SKU order builder */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>ORDER PLACED</label>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <select
                  value=""
                  onChange={(e) => {
                    const sku = SKU_CATALOGUE.find(s => s.id === e.target.value);
                    if (sku) addSku(sku);
                    e.target.value = '';
                  }}
                  style={{ padding: '6px 10px', border: '1px solid rgba(0,53,83,0.25)', background: '#FFFEF2', fontFamily: "'Cinzel', serif", fontSize: 10, letterSpacing: '0.12em', color: '#003553', cursor: 'pointer', fontWeight: 700 }}
                >
                  <option value="">+ ADD SKU</option>
                  {SKU_CATALOGUE.map(sku => (
                    <option key={sku.id} value={sku.id}>
                      {sku.name} — {ZAR(sku.price)}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {items.length === 0 ? (
              <div className="border-2 border-dashed border p-5 text-center bg-cream">
                <p className="text-xs italic ocean">No SKUs added. Tap "ADD SKU" to start building the order.</p>
              </div>
            ) : (
              <div className="border border">
                {/* Mobile: stacked item rows */}
                <div className="md:hidden divide-y divide-ink/10">
                  {items.map(it => {
                    const lineTotal = (Number(it.unitPrice) || 0) * (Number(it.qty) || 0);
                    const discounted = Number(it.unitPrice) < Number(it.listPrice);
                    return (
                      <div key={it.skuId} className="p-3">
                        <div className="flex items-start justify-between gap-2 mb-2">
                          <div>
                            <p className="ink text-sm font-display" style={{ fontWeight: 700 }}>{it.name}</p>
                            {discounted && <p className="text-[10px] copper italic">disc. from {ZAR(it.listPrice)}</p>}
                          </div>
                          <button type="button" onClick={() => removeItem(it.skuId)} className="p-1 text-ink/40 hover:text-red-700 flex-shrink-0">
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                        <div className="grid grid-cols-3 gap-2">
                          <div>
                            <p className="font-display text-[9px] tracking-[0.15em] ocean mb-1" style={{ fontWeight: 600 }}>QTY</p>
                            <input type="number" min="0" step="1" value={it.qty}
                              onChange={(e) => updateItem(it.skuId, 'qty', e.target.value)}
                              className="w-full px-2 py-2 border border bg-cream text-sm text-center focus:outline-none focus:border-copper" />
                          </div>
                          <div>
                            <p className="font-display text-[9px] tracking-[0.15em] ocean mb-1" style={{ fontWeight: 600 }}>UNIT R</p>
                            <input type="number" min="0" step="0.01" value={it.unitPrice}
                              onChange={(e) => updateItem(it.skuId, 'unitPrice', e.target.value)}
                              className="w-full px-2 py-2 border border bg-cream text-sm text-right focus:outline-none focus:border-copper"
                              style={{ color: discounted ? '#D78433' : '#003553' }} />
                          </div>
                          <div>
                            <p className="font-display text-[9px] tracking-[0.15em] ocean mb-1" style={{ fontWeight: 600 }}>TOTAL</p>
                            <p className="font-display text-sm ink py-2 text-right" style={{ fontWeight: 700 }}>{ZAR(lineTotal)}</p>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
                {/* Desktop: compact grid */}
                <div className="hidden md:block">
                  <div className="grid grid-cols-12 gap-2 px-3 py-2  border-b border">
                    <div className="col-span-5 font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>SKU</div>
                    <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-center" style={{ fontWeight: 600 }}>QTY</div>
                    <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-right" style={{ fontWeight: 600 }}>UNIT R</div>
                    <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-right" style={{ fontWeight: 600 }}>LINE</div>
                    <div className="col-span-1"></div>
                  </div>
                  {items.map(it => {
                    const lineTotal = (Number(it.unitPrice) || 0) * (Number(it.qty) || 0);
                    const discounted = Number(it.unitPrice) < Number(it.listPrice);
                    return (
                      <div key={it.skuId} className="grid grid-cols-12 gap-2 px-3 py-2 border-b border items-center">
                        <div className="col-span-5">
                          <p className="ink text-xs font-display" style={{ fontWeight: 700 }}>{it.name}</p>
                          {discounted && <p className="text-[9px] copper italic">from {ZAR(it.listPrice)}</p>}
                        </div>
                        <div className="col-span-2">
                          <input type="number" min="0" step="1" value={it.qty}
                            onChange={(e) => updateItem(it.skuId, 'qty', e.target.value)}
                            className="w-full px-2 py-1 border border bg-cream text-xs text-center focus:outline-none focus:border-copper" />
                        </div>
                        <div className="col-span-2">
                          <input type="number" min="0" step="0.01" value={it.unitPrice}
                            onChange={(e) => updateItem(it.skuId, 'unitPrice', e.target.value)}
                            className="w-full px-2 py-1 border border bg-cream text-xs text-right focus:outline-none focus:border-copper"
                            style={{ color: discounted ? '#D78433' : '#003553' }} />
                        </div>
                        <div className="col-span-2 text-right">
                          <span className="font-display text-xs ink" style={{ fontWeight: 700 }}>{ZAR(lineTotal)}</span>
                        </div>
                        <div className="col-span-1 text-right">
                          <button type="button" onClick={() => removeItem(it.skuId)} className="text-ink/40 hover:text-red-700">
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
                <div className="grid grid-cols-12 gap-2 px-3 py-3" style={{ background: '#003553' }}>
                  <div className="col-span-8 font-display text-[10px] tracking-[0.3em]" style={{ color: '#FDB940', fontWeight: 600 }}>ORDER TOTAL (EX VAT)</div>
                  <div className="col-span-4 text-right font-display text-base" style={{ color: '#FFFEF2', fontWeight: 700 }}>{ZAR(orderTotal)}</div>
                </div>
              </div>
            )}
            <p className="text-[10px] italic ocean mt-1.5">Edit unit price for discounts. Total auto-calculates.</p>
          </div>

          {/* Notes */}
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>VISIT NOTES</label>
            <textarea value={notes} onChange={(e) => setNotes(e.target.value)} rows="3" placeholder="What happened in the meeting? Key decision-makers, objections, opportunities..." className="w-full px-3 py-2 border border bg-cream font-body text-sm focus:outline-none focus:border-copper resize-none" />
          </div>

          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>FOLLOW-UP NOTES</label>
            <textarea value={followUp} onChange={(e) => setFollowUp(e.target.value)} rows="2" placeholder="Next action, due date, who owns it..." className="w-full px-3 py-2 border border bg-cream font-body text-sm focus:outline-none focus:border-copper resize-none" />
          </div>

          {/* Actions */}
          {validationError && (
            <div className="px-3 py-2 border-l-4 text-xs" style={{ borderLeftColor: '#9c2c2c', background: 'rgba(156,44,44,0.08)', color: '#9c2c2c' }}>
              <span className="font-display tracking-wider" style={{ fontWeight: 700 }}>ERROR:</span> {validationError}
            </div>
          )}
          {/* Mobile: stacked full-width buttons. Desktop: row layout */}
          <div className="pt-4 border-t border space-y-2 md:space-y-0 md:flex md:items-center md:justify-between md:gap-3 md:flex-wrap">
            <button type="button" onClick={handleEmailOrder}
              className="w-full md:w-auto flex items-center justify-center gap-2 px-4 py-3 font-display text-xs tracking-[0.25em] border"
              style={{ color: '#006C90', fontWeight: 700, borderColor: '#006C90' }}>
              <Mail className="w-4 h-4" /> EMAIL ORDER
            </button>
            <div className="flex items-center gap-2">
              <button type="button" onClick={onClose} className="flex-1 md:flex-none px-5 py-3 font-display text-xs tracking-[0.25em] ink border border" style={{ fontWeight: 700 }}>CANCEL</button>
              <button type="button" onClick={handleSave} className="flex-1 md:flex-none bg-ink px-6 py-3 font-display text-xs tracking-[0.25em] flex items-center justify-center gap-2" style={{ color: '#FFFEF2', fontWeight: 700 }}>
                {isEdit ? 'UPDATE' : 'SAVE'} <ArrowUpRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>
      </div>

      {emailModalOpen && (
        <EmailRecipientModal
          salesRep={salesRep}
          senderEmail={REP_EMAILS[salesRep] || ''}
          client={selectedClient}
          orderTotal={items.reduce((s, it) => s + (Number(it.unitPrice) || 0) * (Number(it.qty) || 0), 0)}
          itemCount={items.length}
          composeOrderEmail={composeOrderEmail}
          buildMailtoUrl={buildMailtoUrl}
          onClose={() => setEmailModalOpen(false)}
          onSend={(recipient) => {
            sendOrderEmail(recipient);
            setEmailModalOpen(false);
          }}
        />
      )}
    </div>
  );
}

// =================== Email Recipient Modal ===================
// Pop-up that asks the user to confirm/enter the recipient address before
// firing the mailto: link. Shows a preview of the email body so the rep can
// double-check before sending. Defaults to the two standard order-confirmation
// recipients plus the client's email if on file.
function EmailRecipientModal({ salesRep, senderEmail, client, orderTotal, itemCount, composeOrderEmail, buildMailtoUrl, onClose, onSend }) {
  // Build initial recipient list: defaults + client's email if on file (deduped, in order)
  const initialRecipients = useMemo(() => {
    const list = [...DEFAULT_ORDER_RECIPIENTS];
    if (client?.email && !list.includes(client.email)) list.push(client.email);
    return list.join(', ');
  }, [client]);

  const [recipient, setRecipient] = useState(initialRecipients);
  const [showPreview, setShowPreview] = useState(false);
  const [error, setError] = useState('');
  const composed = useMemo(() => composeOrderEmail(), [composeOrderEmail]);

  const isValidEmail = (s) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s.trim());

  // Parse a comma- (or semicolon-) separated list of addresses
  const parseRecipients = (raw) =>
    raw.split(/[,;]/).map(s => s.trim()).filter(Boolean);

  const send = () => {
    const list = parseRecipients(recipient);
    if (list.length === 0) {
      setError('Please enter at least one recipient email address.');
      return;
    }
    const invalid = list.filter(e => !isValidEmail(e));
    if (invalid.length > 0) {
      setError(`Invalid email${invalid.length > 1 ? 's' : ''}: ${invalid.join(', ')}`);
      return;
    }
    setError('');
    // Join with comma — mailto supports multiple recipients separated by commas
    onSend(list.join(','));
  };

  // Quick toggle: include/exclude client email
  const includesClient = client?.email && parseRecipients(recipient).includes(client.email);
  const toggleClientEmail = () => {
    const list = parseRecipients(recipient);
    if (includesClient) {
      setRecipient(list.filter(e => e !== client.email).join(', '));
    } else {
      setRecipient([...list, client.email].join(', '));
    }
  };

  // Reset to defaults
  const resetToDefaults = () => {
    setRecipient(initialRecipients);
    setError('');
  };

  return (
    <div style={{ position:"fixed", inset:0, zIndex:70, display:"flex", alignItems:"center", justifyContent:"center", padding:16, background:"rgba(0,53,83,0.82)", backdropFilter:"blur(4px)", overflowY:"auto" }} onClick={onClose}>
      <div className="bg-cream max-w-2xl w-full my-4 border-2 border-copper" onClick={(e) => e.stopPropagation()}>
        <div className="p-5 flex items-center justify-between gap-4" style={{ background: '#003553' }}>
          <div className="flex items-center gap-3">
            <Mail className="w-5 h-5" style={{ color: '#FDB940' }} />
            <div>
              <p className="font-display text-[10px] tracking-[0.4em]" style={{ color: '#FDB940', fontWeight: 600 }}>SEND ORDER CONFIRMATION</p>
              <h2 className="font-display text-xl mt-0.5" style={{ color: '#FFFEF2', fontWeight: 700, letterSpacing: '0.06em' }}>EMAIL ORDER</h2>
            </div>
          </div>
          <button type="button" onClick={onClose} className="p-2 hover:opacity-70" style={{ color: '#FFFEF2' }}><X className="w-5 h-5" /></button>
        </div>

        <div className="p-6 space-y-5">
          {/* From */}
          <div>
            <p className="font-display text-[10px] tracking-[0.3em] copper mb-1" style={{ fontWeight: 600 }}>FROM</p>
            <div className="flex items-center gap-3 p-3  border border">
              <div className="w-8 h-8 flex items-center justify-center bg-ink">
                <span className="font-display text-sm" style={{ color: '#FDB940', fontWeight: 700 }}>{salesRep[0]}</span>
              </div>
              <div className="min-w-0">
                <p className="font-display text-sm ink" style={{ fontWeight: 700 }}>{salesRep} · Avante Cape Brandy</p>
                <p className="text-xs ocean italic truncate">{senderEmail || '(no email on file for this rep)'}</p>
              </div>
            </div>
          </div>

          {/* Recipient */}
          <div>
            <div className="flex items-baseline justify-between mb-1 flex-wrap gap-2">
              <p className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>SEND TO</p>
              <div className="flex items-center gap-3">
                {client?.email && (
                  <button type="button" onClick={toggleClientEmail} className="text-[10px] font-display tracking-wider copper hover:gold" style={{ fontWeight: 700 }}>
                    {includesClient ? '− REMOVE CLIENT EMAIL' : '+ ADD CLIENT EMAIL'}
                  </button>
                )}
                <button type="button" onClick={resetToDefaults} className="text-[10px] font-display tracking-wider ocean hover:ink" style={{ fontWeight: 700 }}>
                  RESET
                </button>
              </div>
            </div>
            <textarea
              value={recipient}
              onChange={(e) => { setRecipient(e.target.value); if (error) setError(''); }}
              placeholder="recipient@example.com, another@example.com"
              rows="2"
              className="w-full px-3 py-2 border border bg-cream font-body text-sm focus:outline-none focus:border-copper resize-none"
              autoFocus
            />
            <p className="text-[10px] italic ocean mt-1">Separate multiple addresses with commas. Defaults: {DEFAULT_ORDER_RECIPIENTS.join(', ')}.</p>
            {client?.venue && (
              <p className="text-[11px] italic ocean mt-1.5">
                Order for <strong className="ink">{client.venue}</strong>
                {client.firstName || client.lastName ? ` · attn. ${[client.firstName, client.lastName].filter(Boolean).join(' ')}` : ''}
              </p>
            )}
            {error && (
              <div className="mt-2 px-3 py-2 border-l-4 text-xs" style={{ borderLeftColor: '#9c2c2c', background: 'rgba(156,44,44,0.08)', color: '#9c2c2c' }}>
                <span className="font-display tracking-wider" style={{ fontWeight: 700 }}>ERROR:</span> {error}
              </div>
            )}
          </div>

          {/* Subject summary */}
          <div>
            <p className="font-display text-[10px] tracking-[0.3em] copper mb-1" style={{ fontWeight: 600 }}>SUBJECT</p>
            <p className="text-xs ink p-3  border border italic">{composed.subject}</p>
          </div>

          {/* Order summary */}
          <div className="grid grid-cols-3 gap-3 py-3 border-y border">
            <div>
              <p className="font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>SKUs</p>
              <p className="font-display text-lg ink" style={{ fontWeight: 700 }}>{itemCount}</p>
            </div>
            <div>
              <p className="font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>ORDER TOTAL</p>
              <p className="font-display text-lg copper" style={{ fontWeight: 700 }}>{ZAR(orderTotal)}</p>
            </div>
            <div>
              <p className="font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>VAT</p>
              <p className="font-display text-lg ink" style={{ fontWeight: 700 }}>Excl.</p>
            </div>
          </div>

          {/* Preview toggle */}
          <div>
            <button
              type="button"
              onClick={() => setShowPreview(!showPreview)}
              className="text-[10px] font-display tracking-[0.2em] copper hover:gold"
              style={{ fontWeight: 700 }}
            >
              {showPreview ? '▾ HIDE EMAIL PREVIEW' : '▸ SHOW EMAIL PREVIEW'}
            </button>
            {showPreview && (
              <pre className="mt-2 p-3  border border text-[11px] ink whitespace-pre-wrap font-mono max-h-72 overflow-y-auto scrollbar-thin">{composed.body}</pre>
            )}
          </div>

          <div>
            <p className="font-display text-[10px] tracking-[0.3em] copper mb-2" style={{ fontWeight: 600 }}>SEND VIA</p>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-2">
              {/* Gmail web compose */}
              <a
                href={(() => {
                  const list = parseRecipients(recipient).filter(e => isValidEmail(e));
                  if (list.length === 0) return '#';
                  const to = encodeURIComponent(list.join(','));
                  const su = encodeURIComponent(composed.subject);
                  const bo = encodeURIComponent(composed.body);
                  return `https://mail.google.com/mail/?view=cm&fs=1&to=${to}&su=${su}&body=${bo}`;
                })()}
                onClick={(e) => {
                  const list = parseRecipients(recipient);
                  if (list.length === 0) { e.preventDefault(); setError('Please enter at least one recipient.'); return; }
                  const invalid = list.filter(em => !isValidEmail(em));
                  if (invalid.length > 0) { e.preventDefault(); setError(`Invalid: ${invalid.join(', ')}`); return; }
                  setError('');
                }}
                target="_blank"
                rel="noopener noreferrer"
                className="flex items-center justify-center gap-2 px-4 py-3 font-display text-xs tracking-[0.2em] hover:opacity-90 no-underline border border"
                style={{ background: '#FFFEF2', color: '#003553', fontWeight: 700, textDecoration: 'none' }}
              >
                <Mail className="w-4 h-4" /> GMAIL
              </a>
              {/* Outlook web compose */}
              <a
                href={(() => {
                  const list = parseRecipients(recipient).filter(e => isValidEmail(e));
                  if (list.length === 0) return '#';
                  const to = encodeURIComponent(list.join(';'));
                  const su = encodeURIComponent(composed.subject);
                  const bo = encodeURIComponent(composed.body);
                  return `https://outlook.office.com/mail/deeplink/compose?to=${to}&subject=${su}&body=${bo}`;
                })()}
                onClick={(e) => {
                  const list = parseRecipients(recipient);
                  if (list.length === 0) { e.preventDefault(); setError('Please enter at least one recipient.'); return; }
                  const invalid = list.filter(em => !isValidEmail(em));
                  if (invalid.length > 0) { e.preventDefault(); setError(`Invalid: ${invalid.join(', ')}`); return; }
                  setError('');
                }}
                target="_blank"
                rel="noopener noreferrer"
                className="flex items-center justify-center gap-2 px-4 py-3 font-display text-xs tracking-[0.2em] hover:opacity-90 no-underline border border"
                style={{ background: '#FFFEF2', color: '#003553', fontWeight: 700, textDecoration: 'none' }}
              >
                <Mail className="w-4 h-4" /> OUTLOOK
              </a>
              {/* OS default mail app via mailto */}
              <a
                href={(() => {
                  const list = parseRecipients(recipient);
                  if (list.length === 0 || list.some(e => !isValidEmail(e))) return '#';
                  return buildMailtoUrl(list.join(','));
                })()}
                onClick={(e) => {
                  const list = parseRecipients(recipient);
                  if (list.length === 0) { e.preventDefault(); setError('Please enter at least one recipient.'); return; }
                  const invalid = list.filter(em => !isValidEmail(em));
                  if (invalid.length > 0) { e.preventDefault(); setError(`Invalid: ${invalid.join(', ')}`); return; }
                  setError('');
                }}
                target="_top"
                className="flex items-center justify-center gap-2 px-4 py-3 font-display text-xs tracking-[0.2em] hover:opacity-90 no-underline"
                style={{ background: '#006C90', color: '#FFFEF2', fontWeight: 700, textDecoration: 'none' }}
              >
                <Mail className="w-4 h-4" /> OS MAIL APP
              </a>
            </div>
            <p className="text-[10px] italic ocean mt-2">Pick the one your team uses. Gmail and Outlook open in a new browser tab — no setup required.</p>
          </div>
          {client?.email && !includesClient && (
            <p className="text-[11px] italic" style={{ color: '#D78433' }}>
              Tip: Click "+ ADD CLIENT EMAIL" above to also send to <strong>{client.email}</strong>.
            </p>
          )}
        </div>

        <div className="flex flex-wrap justify-between gap-3 px-6 pb-6">
          <CopyEmailButton recipient={recipient} buildMailtoUrl={buildMailtoUrl} composed={composed} />
          <button type="button" onClick={onClose} className="px-5 py-2.5 font-display text-xs tracking-[0.25em] ink hover:" style={{ fontWeight: 700 }}>
            DONE
          </button>
        </div>
      </div>
    </div>
  );
}

// Helper: copy email body to clipboard as a fallback if mailto doesn't work.
function CopyEmailButton({ recipient, buildMailtoUrl, composed }) {
  const [copied, setCopied] = useState(false);
  const copy = async () => {
    const fullText = `To: ${recipient}\nSubject: ${composed.subject}\n\n${composed.body}`;
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(fullText);
      } else {
        // Fallback: temp textarea
        const ta = document.createElement('textarea');
        ta.value = fullText;
        ta.style.position = 'fixed';
        ta.style.opacity = '0';
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
      }
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error('[CopyEmail] failed', err);
    }
  };
  return (
    <button
      type="button"
      onClick={copy}
      title="Copy the full email (recipients, subject, body) to your clipboard so you can paste into any mail app"
      className="flex items-center gap-2 px-4 py-2.5 font-display text-[10px] tracking-[0.25em] border border hover:"
      style={{ color: '#003553', fontWeight: 700 }}
    >
      {copied ? '✓ COPIED TO CLIPBOARD' : 'COPY EMAIL TEXT'}
    </button>
  );
}

// =================== New Client Modal ===================
function NewClientModal({ defaultRep, onClose, onSave }) {
  const [form, setForm] = useState({
    venue: '',
    firstName: '',
    lastName: '',
    location: '',
    distributor: '',
    email: '',
    phone: '',
    channel: 'Private Sales',
    leadSource: 'Cold Call',
    priority: 'Medium',
    accountManager: defaultRep,
    status: 'New',
    notes: '',
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');

  const setF = (k, v) => { setForm(f => ({ ...f, [k]: v })); setError(''); };

  const handleCreate = async () => {
    // Validate required fields
    if (!form.venue.trim()) { setError('Venue name is required.'); return; }
    if (!form.channel) { setError('Channel is required (Private Sales or Trade Retail).'); return; }
    if (!form.accountManager) { setError('Please assign an account manager.'); return; }

    setSaving(true);
    setError('');
    try {
      const created = await onSave(form);
      // onSave handles closing the modal — if it returns null, something went wrong
      if (!created) {
        setError('Save failed — could not write to database. Check your connection and try again.');
      }
    } catch (err) {
      console.error('[NewClient] save threw', err);
      setError('Save failed: ' + (err?.message || String(err)));
    } finally {
      setSaving(false);
    }
  };

  return (
    <div style={{ position:"fixed", inset:0, zIndex:60, display:"flex", alignItems:"center", justifyContent:"center", padding:16, background:"rgba(0,53,83,0.82)", backdropFilter:"blur(4px)", overflowY:"auto" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background:"#FFFEF2", width:"100%", maxWidth:672, maxHeight:"92vh", overflowY:"auto", border:"2px solid #D78433" }}>

        {/* Header */}
        <div className="p-4 flex items-center justify-between sticky top-0 z-10" style={{ background: '#003553' }}>
          <div className="flex items-center gap-3">
            <UserPlus className="w-5 h-5 flex-shrink-0" style={{ color: '#FDB940' }} />
            <div>
              <p className="font-display text-[9px] tracking-[0.3em]" style={{ color: '#FDB940', fontWeight: 600 }}>NEW CLIENT</p>
              <h2 className="font-display text-lg" style={{ color: '#FFFEF2', fontWeight: 700, letterSpacing: '0.06em' }}>ADD TO {defaultRep.toUpperCase()}'S BOOK</h2>
            </div>
          </div>
          <button type="button" onClick={onClose} className="p-2 hover:opacity-70 flex-shrink-0" style={{ color: '#FFFEF2' }}><X className="w-5 h-5" /></button>
        </div>

        <div className="p-4 space-y-4">
          {/* Venue name — most important field */}
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>VENUE NAME *</label>
            <input
              type="text" value={form.venue}
              onChange={(e) => setF('venue', e.target.value)}
              placeholder="e.g. The Grand Hotel"
              className="w-full px-3 py-3 border-2 bg-cream font-body text-sm focus:outline-none"
              style={{ borderColor: !form.venue.trim() && error ? '#9c2c2c' : 'rgba(0,53,83,0.2)' }}
              autoFocus
            />
          </div>

          {/* Channel — critical for visit logging */}
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>CHANNEL * <span className="italic normal-case text-[9px] ocean">(required for visit logging)</span></label>
            <div className="grid grid-cols-2 gap-2">
              {CHANNELS.map(ch => (
                <button key={ch} type="button"
                  onClick={() => setF('channel', ch)}
                  className={`py-3 font-display text-xs tracking-[0.15em] border-2 transition-all ${form.channel === ch ? 'bg-ink border-ink' : 'bg-cream border'}`}
                  style={{ color: form.channel === ch ? '#FFFEF2' : '#003553', fontWeight: 700 }}>
                  {ch === 'Private Sales' ? 'PRIVATE SALES' : 'TRADE RETAIL'}
                </button>
              ))}
            </div>
          </div>

          {/* Two-column fields */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="font-display text-[10px] tracking-[0.25em] copper mb-1 block" style={{ fontWeight: 600 }}>FIRST NAME</label>
              <input type="text" value={form.firstName} onChange={(e) => setF('firstName', e.target.value)} className="w-full px-3 py-2.5 border border bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.25em] copper mb-1 block" style={{ fontWeight: 600 }}>LAST NAME</label>
              <input type="text" value={form.lastName} onChange={(e) => setF('lastName', e.target.value)} className="w-full px-3 py-2.5 border border bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.25em] copper mb-1 block" style={{ fontWeight: 600 }}>EMAIL</label>
              <input type="email" value={form.email} onChange={(e) => setF('email', e.target.value)} className="w-full px-3 py-2.5 border border bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.25em] copper mb-1 block" style={{ fontWeight: 600 }}>PHONE</label>
              <input type="tel" value={form.phone} onChange={(e) => setF('phone', e.target.value)} className="w-full px-3 py-2.5 border border bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.25em] copper mb-1 block" style={{ fontWeight: 600 }}>LOCATION</label>
              <input type="text" value={form.location} onChange={(e) => setF('location', e.target.value)} placeholder="e.g. Camps Bay" className="w-full px-3 py-2.5 border border bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.25em] copper mb-1 block" style={{ fontWeight: 600 }}>DISTRIBUTOR</label>
              <input type="text" value={form.distributor} onChange={(e) => setF('distributor', e.target.value)} className="w-full px-3 py-2.5 border border bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.25em] copper mb-1 block" style={{ fontWeight: 600 }}>LEAD SOURCE</label>
              <select value={form.leadSource} onChange={(e) => setF('leadSource', e.target.value)} className="w-full px-3 py-2.5 border border bg-cream font-body text-sm focus:outline-none focus:border-copper">
                {LEAD_SOURCES.map(s => <option key={s}>{s}</option>)}
              </select>
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.25em] copper mb-1 block" style={{ fontWeight: 600 }}>PRIORITY</label>
              <select value={form.priority} onChange={(e) => setF('priority', e.target.value)} className="w-full px-3 py-2.5 border border bg-cream font-body text-sm focus:outline-none focus:border-copper">
                {PRIORITIES.map(p => <option key={p}>{p}</option>)}
              </select>
            </div>
          </div>

          {/* Account manager — full width */}
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>ACCOUNT MANAGER *</label>
            <div className="grid grid-cols-3 gap-2">
              {SALES_REPS.map(r => (
                <button key={r} type="button"
                  onClick={() => setF('accountManager', r)}
                  className={`py-2.5 font-display text-xs tracking-[0.15em] border transition-all ${form.accountManager === r ? 'bg-ink border-ink' : 'bg-cream border'}`}
                  style={{ color: form.accountManager === r ? '#FFFEF2' : '#003553', fontWeight: 700 }}>
                  {r.toUpperCase()}
                </button>
              ))}
            </div>
          </div>

          {/* Notes */}
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>NOTES</label>
            <textarea value={form.notes} onChange={(e) => setF('notes', e.target.value)} rows="2" placeholder="Any initial context about this venue..." className="w-full px-3 py-2 border border bg-cream font-body text-sm focus:outline-none focus:border-copper resize-none" />
          </div>

          {/* Error display */}
          {error && (
            <div className="px-3 py-3 border-l-4 text-xs" style={{ borderLeftColor: '#9c2c2c', background: 'rgba(156,44,44,0.08)', color: '#9c2c2c' }}>
              <span className="font-display tracking-wider" style={{ fontWeight: 700 }}>ERROR: </span>{error}
            </div>
          )}

          {/* Confirmation message */}
          <p className="text-[10px] italic ocean">
            This client will be saved to Supabase and immediately available for visit logging.
          </p>

          {/* Actions */}
          <div className="flex gap-3 pt-3 border-t border">
            <button type="button" onClick={onClose} className="flex-none px-4 py-3 font-display text-xs tracking-[0.2em] ink border border" style={{ fontWeight: 700 }}>
              CANCEL
            </button>
            <button
              type="button"
              onClick={handleCreate}
              disabled={saving}
              className={`flex-1 flex items-center justify-center gap-2 py-3 font-display text-xs tracking-[0.2em] transition-all ${saving ? 'opacity-60 cursor-not-allowed' : ''}`}
              style={{ background: '#D78433', color: '#FFFEF2', fontWeight: 700 }}
            >
              {saving ? (
                <>
                  <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>
                  SAVING TO DATABASE...
                </>
              ) : (
                <><UserPlus className="w-4 h-4" /> CREATE &amp; ADD TO BOOK</>
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// =================== Client Detail Modal ===================
function ClientDetailModal({ client, visits, onClose, onUpdate }) {
  const [edit, setEdit] = useState(false);
  const [form, setForm] = useState(client);

  const save = () => {
    console.log('[ClientDetail] saving', form);
    try {
      onUpdate(form);
      setEdit(false);
    } catch (err) {
      console.error('[ClientDetail] save threw:', err);
      alert('Could not save client: ' + (err?.message || err));
    }
  };

  // Build the visit timeline. We merge two sources:
  // 1) Visits logged through the CRM (the `visits` prop)
  // 2) The original "Last Contacted" date from the imported spreadsheet
  //    — surfaced as a historical entry so reps can see prior contact history
  //    even before the CRM was in use. Only added if no logged visit exists
  //    on that exact date (to avoid duplication once a real visit is logged).
  const orderedVisits = useMemo(() => {
    const list = visits.slice();
    const seedDate = client.lastContacted;
    const hasLoggedOnSeedDate = seedDate && list.some(v => v.date === seedDate);
    if (seedDate && !hasLoggedOnSeedDate) {
      list.push({
        id: `seed-${client.id}`,
        isHistorical: true, // flag so we can render differently
        date: seedDate,
        salesRep: client.accountManager && client.accountManager !== 'Unassigned' ? client.accountManager : '—',
        outcome: 'Prior Contact',
        saleAmount: 0,
        items: [],
        notes: client.notes || 'Imported from prior records — no detailed log available.',
        followUp: '',
      });
    }
    return list.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  }, [visits, client.lastContacted, client.id, client.accountManager, client.notes]);

  // Aggregate SKU purchases across all visits
  const skuTotals = useMemo(() => {
    const tally = new Map();
    visits.forEach(v => {
      (v.items || []).forEach(it => {
        const cur = tally.get(it.skuId) || { skuId: it.skuId, name: it.name, qty: 0, revenue: 0, lastDate: '' };
        cur.qty += Number(it.qty) || 0;
        cur.revenue += Number(it.lineTotal) || (Number(it.qty || 0) * Number(it.unitPrice || 0));
        if (!cur.lastDate || (v.date || '') > cur.lastDate) cur.lastDate = v.date || '';
        tally.set(it.skuId, cur);
      });
    });
    return Array.from(tally.values()).sort((a, b) => b.revenue - a.revenue);
  }, [visits]);

  // Days since last contact
  const daysSinceLast = useMemo(() => {
    if (!client.lastContacted) return null;
    const last = new Date(client.lastContacted);
    if (isNaN(last)) return null;
    const diff = Math.floor((Date.now() - last.getTime()) / (1000 * 60 * 60 * 24));
    return diff;
  }, [client.lastContacted]);

  return (
    <div style={{ position:"fixed", inset:0, zIndex:50, display:"flex", alignItems:"center", justifyContent:"center", padding:16, background:"rgba(0,53,83,0.82)", backdropFilter:"blur(4px)", overflowY:"auto" }} onClick={onClose}>
      <div className="bg-cream w-full md:max-w-3xl md:my-4 border-t-2 md:border-2 border-copper max-h-[95vh] md:max-h-none overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="p-4 md:p-6 relative overflow-hidden" style={{ background: '#003553' }}>
          <RaysBackdrop opacity={0.06} />
          <div className="flex items-start justify-between gap-4 relative">
            <div className="flex items-start gap-3">
              <AvanteLogo height={36} />
              <div>
                <p className="font-display text-[9px] tracking-[0.3em]" style={{ color: '#FDB940', fontWeight: 600 }}>{client.channel?.toUpperCase() || 'NO CHANNEL'}</p>
                <h2 className="font-display text-xl md:text-3xl mt-0.5" style={{ color: '#FFFEF2', fontWeight: 700, letterSpacing: '0.05em' }}>{client.venue}</h2>
                <p className="italic text-xs mt-0.5" style={{ color: '#FDB940' }}>{client.location || 'No location'}</p>
              </div>
            </div>
            <button type="button" onClick={onClose} className="p-2 hover:opacity-70" style={{ color: '#FFFEF2' }}><X className="w-5 h-5" /></button>
          </div>
        </div>

        <div className="p-4 md:p-6 space-y-5">
          {/* Quick stats — 2 col on mobile, 4 on desktop */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 pb-4 border-b border">
            <div>
              <p className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>TOTAL SALES TD</p>
              <p className="font-display text-xl ink mt-1" style={{ fontWeight: 700 }}>{ZAR(client.totalSales || 0)}</p>
            </div>
            <div>
              <p className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>VISITS LOGGED</p>
              <p className="font-display text-xl ink mt-1" style={{ fontWeight: 700 }}>{visits.length}</p>
            </div>
            <div>
              <p className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>LAST VISIT</p>
              <p className="font-display text-base ink mt-1" style={{ fontWeight: 700 }}>{client.lastContacted || '—'}</p>
              {daysSinceLast !== null && (
                <p className="text-[10px] italic ocean mt-0.5">{daysSinceLast === 0 ? 'today' : daysSinceLast === 1 ? '1 day ago' : `${daysSinceLast} days ago`}</p>
              )}
            </div>
            <div>
              <p className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>STATUS</p>
              <div className="mt-2"><StatusBadge status={client.status} /></div>
            </div>
          </div>

          {/* Contact / classification fields */}
          <div>
            <p className="font-display text-[10px] tracking-[0.3em] copper mb-3" style={{ fontWeight: 600 }}>CONTACT &amp; CLASSIFICATION</p>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <Field label="First Name" value={form.firstName} edit={edit} onChange={(v) => setForm({ ...form, firstName: v })} />
              <Field label="Last Name" value={form.lastName} edit={edit} onChange={(v) => setForm({ ...form, lastName: v })} />
              <Field label="Email" value={form.email} edit={edit} onChange={(v) => setForm({ ...form, email: v })} icon={Mail} />
              <Field label="Phone" value={form.phone} edit={edit} onChange={(v) => setForm({ ...form, phone: v })} icon={Phone} />
              <Field label="Distributor" value={form.distributor} edit={edit} onChange={(v) => setForm({ ...form, distributor: v })} />
              <SelectField label="Account Manager" value={form.accountManager} edit={edit} onChange={(v) => setForm({ ...form, accountManager: v })} options={[...SALES_REPS, 'Unassigned']} />
              <SelectField label="Lead Source" value={form.leadSource} edit={edit} onChange={(v) => setForm({ ...form, leadSource: v })} options={LEAD_SOURCES} />
              <SelectField label="Priority" value={form.priority} edit={edit} onChange={(v) => setForm({ ...form, priority: v })} options={PRIORITIES} />
              <SelectField label="Status" value={form.status} edit={edit} onChange={(v) => setForm({ ...form, status: v })} options={STATUSES} />
              <SelectField label="Channel" value={form.channel} edit={edit} onChange={(v) => setForm({ ...form, channel: v })} options={CHANNELS} />
            </div>
          </div>

          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>NOTES</label>
            {edit ? (
              <textarea value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} rows="3" className="w-full px-3 py-2 border border bg-cream font-body text-sm focus:outline-none focus:border-copper resize-none" />
            ) : (
              <p className="text-sm ink italic">{client.notes || '—'}</p>
            )}
          </div>

          {/* === PURCHASE HISTORY (SKUs aggregated) === */}
          {skuTotals.length > 0 && (
            <div className="pt-5 border-t border">
              <div className="flex items-center justify-between mb-3">
                <p className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>PURCHASE HISTORY · SKU BREAKDOWN</p>
                <span className="text-[10px] italic ocean">All-time totals across {visits.filter(v => v.items && v.items.length > 0).length} order{visits.filter(v => v.items && v.items.length > 0).length === 1 ? '' : 's'}</span>
              </div>
              <div className="border border">
                <div className="grid grid-cols-12 gap-2 px-3 py-2  border-b border">
                  <div className="col-span-6 font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>SKU</div>
                  <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-center" style={{ fontWeight: 600 }}>UNITS</div>
                  <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-right" style={{ fontWeight: 600 }}>REVENUE</div>
                  <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-right" style={{ fontWeight: 600 }}>LAST</div>
                </div>
                {skuTotals.map(s => (
                  <div key={s.skuId} className="grid grid-cols-12 gap-2 px-3 py-2 border-b border items-center text-xs">
                    <div className="col-span-6 ink font-display" style={{ fontWeight: 700 }}>{s.name}</div>
                    <div className="col-span-2 text-center ink">{s.qty}</div>
                    <div className="col-span-2 text-right copper font-display" style={{ fontWeight: 700 }}>{ZAR(s.revenue)}</div>
                    <div className="col-span-2 text-right ocean italic">{s.lastDate || '—'}</div>
                  </div>
                ))}
                <div className="grid grid-cols-12 gap-2 px-3 py-2.5" style={{ background: '#003553' }}>
                  <div className="col-span-6 font-display text-[10px] tracking-[0.3em]" style={{ color: '#FDB940', fontWeight: 600 }}>LIFETIME TOTAL</div>
                  <div className="col-span-2 text-center font-display" style={{ color: '#FFFEF2', fontWeight: 700 }}>{skuTotals.reduce((s, x) => s + x.qty, 0)}</div>
                  <div className="col-span-4 text-right font-display text-base" style={{ color: '#FFFEF2', fontWeight: 700 }}>{ZAR(skuTotals.reduce((s, x) => s + x.revenue, 0))}</div>
                </div>
              </div>
            </div>
          )}

          {/* === CALL CYCLE / VISIT TIMELINE === */}
          {orderedVisits.length > 0 ? (
            <div className="pt-5 border-t border">
              <div className="flex items-center justify-between mb-3">
                <p className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>CALL CYCLE · VISIT TIMELINE ({orderedVisits.length})</p>
                <span className="text-[10px] italic ocean">Newest first</span>
              </div>
              <div className="space-y-3">
                {orderedVisits.map((v, idx) => {
                  const isHistorical = v.isHistorical;
                  // Historical entries get a muted ocean blue; logged visits get outcome-based color
                  const outcomeColor = isHistorical
                    ? '#006C90'
                    : v.outcome === 'Sold In' ? '#2d8659'
                    : v.outcome === 'Rejected' ? '#9c2c2c'
                    : '#D78433';
                  const isLatest = idx === 0 && !isHistorical;
                  return (
                    <div key={v.id} className="relative pl-6">
                      {/* Timeline dot */}
                      <div
                        className="absolute left-0 top-2 w-3 h-3 diamond-clip"
                        style={{ background: outcomeColor, opacity: isHistorical ? 0.6 : 1 }}
                      ></div>
                      {/* Connecting line */}
                      {idx < orderedVisits.length - 1 && (
                        <div className="absolute left-[5px] top-5 bottom-[-12px] w-px bg-ink/15"></div>
                      )}
                      <div className={`border ${isLatest ? 'border-copper' : isHistorical ? 'border border-dashed' : 'border'} bg-cream`}>
                        <div className="flex items-baseline justify-between gap-2 px-3 py-2 border-b border" style={{ background: isLatest ? 'rgba(253,185,64,0.08)' : isHistorical ? 'rgba(0,108,144,0.04)' : 'transparent' }}>
                          <div className="flex items-baseline gap-3 flex-wrap">
                            <span className="font-display ink text-sm tracking-wide" style={{ fontWeight: 700 }}>{v.date || 'No date'}</span>
                            <span className="text-[10px] font-display tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>BY {v.salesRep?.toUpperCase() || '—'}</span>
                            {isLatest && <span className="text-[9px] font-display tracking-[0.2em] px-1.5 py-0.5" style={{ background: '#FDB940', color: '#003553', fontWeight: 700 }}>LATEST</span>}
                            {isHistorical && <span className="text-[9px] font-display tracking-[0.2em] px-1.5 py-0.5 border" style={{ borderColor: '#006C90', color: '#006C90', fontWeight: 700 }}>IMPORTED</span>}
                          </div>
                          <div className="flex items-center gap-3">
                            <span className="text-[10px] font-display tracking-wider px-2 py-0.5" style={{ background: outcomeColor, color: '#FFFEF2', fontWeight: 600 }}>
                              {(v.outcome || 'VISIT').toUpperCase()}
                            </span>
                            {v.saleAmount > 0 && <span className="font-display text-sm copper" style={{ fontWeight: 700 }}>{ZAR(v.saleAmount)}</span>}
                          </div>
                        </div>
                        <div className="px-3 py-2 space-y-2">
                          {/* SKU items */}
                          {v.items && v.items.length > 0 && (
                            <div>
                              <p className="font-display text-[9px] tracking-[0.2em] ocean mb-1" style={{ fontWeight: 600 }}>ORDER · {v.items.length} {v.items.length === 1 ? 'SKU' : 'SKUS'}</p>
                              <div className="grid grid-cols-1 md:grid-cols-2 gap-1">
                                {v.items.map((it, i) => {
                                  const discounted = Number(it.unitPrice) < Number(it.listPrice);
                                  return (
                                    <div key={i} className="flex items-center justify-between text-[11px] px-2 py-1 ">
                                      <span className="ink">
                                        <span className="font-display tracking-wider" style={{ fontWeight: 700 }}>{it.qty}×</span> {it.name}
                                        {discounted && <span className="copper italic"> · disc.</span>}
                                      </span>
                                      <span className="ocean">{ZAR(it.lineTotal)}</span>
                                    </div>
                                  );
                                })}
                              </div>
                            </div>
                          )}
                          {/* Legacy free-text order (pre-SKU upgrade) */}
                          {v.orderPlaced && (!v.items || v.items.length === 0) && (
                            <div>
                              <p className="font-display text-[9px] tracking-[0.2em] ocean mb-0.5" style={{ fontWeight: 600 }}>ORDER PLACED</p>
                              <p className="text-xs ink italic">{v.orderPlaced}</p>
                            </div>
                          )}
                          {/* Notes */}
                          {v.notes && (
                            <div>
                              <p className="font-display text-[9px] tracking-[0.2em] ocean mb-0.5" style={{ fontWeight: 600 }}>{isHistorical ? 'CONTEXT' : 'VISIT NOTES'}</p>
                              <p className="text-xs ink italic">{v.notes}</p>
                            </div>
                          )}
                          {/* Follow-up */}
                          {v.followUp && (
                            <div>
                              <p className="font-display text-[9px] tracking-[0.2em] copper mb-0.5" style={{ fontWeight: 600 }}>FOLLOW-UP</p>
                              <p className="text-xs ink italic">{v.followUp}</p>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ) : (
            <div className="pt-5 border-t border">
              <p className="font-display text-[10px] tracking-[0.3em] copper mb-3" style={{ fontWeight: 600 }}>CALL CYCLE · VISIT TIMELINE</p>
              <div className="border border-dashed border p-6 text-center">
                <p className="text-xs italic ocean">No visits logged for this client yet.</p>
                <p className="text-xs italic ocean mt-1">Use "Log Visit" to start tracking the call cycle.</p>
              </div>
            </div>
          )}

          <div className="flex items-center justify-end gap-3 pt-4 border-t border">
            {edit ? (
              <>
                <button type="button" onClick={() => { setForm(client); setEdit(false); }} className="px-5 py-2.5 font-display text-xs tracking-[0.25em] ink" style={{ fontWeight: 700 }}>CANCEL</button>
                <button type="button" onClick={save} className="bg-copper hover:bg-gold px-6 py-2.5 font-display text-xs tracking-[0.25em]" style={{ color: '#FFFEF2', fontWeight: 700 }}>SAVE CHANGES</button>
              </>
            ) : (
              <button type="button" onClick={() => setEdit(true)} className="bg-ink hover:bg-copper px-6 py-2.5 font-display text-xs tracking-[0.25em]" style={{ color: '#FFFEF2', fontWeight: 700 }}>EDIT CLIENT</button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

function Field({ label, value, edit, onChange, icon: Icon }) {
  return (
    <div>
      <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>{label.toUpperCase()}</label>
      {edit ? (
        <input type="text" value={value || ''} onChange={(e) => onChange(e.target.value)} className="w-full px-3 py-2 border border bg-cream font-body text-sm focus:outline-none focus:border-copper" />
      ) : (
        <div className="flex items-center gap-2 text-sm ink py-1">
          {Icon && <Icon className="w-3.5 h-3.5 ocean" />}
          {value || <span className="italic ocean">—</span>}
        </div>
      )}
    </div>
  );
}

function SelectField({ label, value, edit, onChange, options }) {
  return (
    <div>
      <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>{label.toUpperCase()}</label>
      {edit ? (
        <select value={value || ''} onChange={(e) => onChange(e.target.value)} className="w-full px-3 py-2 border border bg-cream font-body text-sm focus:outline-none focus:border-copper">
          {options.map(o => <option key={o} value={o}>{o}</option>)}
        </select>
      ) : (
        <div className="text-sm ink py-1">{value || <span className="italic ocean">—</span>}</div>
      )}
    </div>
  );
}
