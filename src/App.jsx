import React, { useState, useEffect, useMemo, useRef } from 'react';
import { LayoutDashboard, Users, ClipboardList, Settings, TrendingUp, Phone, Mail, Search, Plus, X, ChevronRight, DollarSign, Award, Activity, Briefcase, Wine, ArrowUpRight, Save, RotateCcw, Target, BarChart3, Trash2, Download, FileSpreadsheet, UserPlus, Edit2 } from 'lucide-react';
import * as XLSX from 'xlsx';

const SEED_CLIENTS = [{"id":1,"venue":"The Lawns","channel":"Private Sales","firstName":"","lastName":"","location":"Camps Bay","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Low","lastContacted":"","status":"New","notes":"Need to make contact","totalSales":0},{"id":2,"venue":"50 on Gugs","channel":"Private Sales","firstName":"Sinbad","lastName":"","location":"Gugulethu","distributor":"Ultra","email":"","phone":"082 553 3599","leadSource":"Referral","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-18","status":"New","notes":"","totalSales":0},{"id":3,"venue":"Mitchell's","channel":"Private Sales","firstName":"Leroy","lastName":"","location":"CBD","distributor":"","email":"","phone":"660793944","leadSource":"Walk into Store","accountManager":"Alex","priority":"Medium","lastContacted":"2026-01-29","status":"New","notes":"Parking for now","totalSales":0},{"id":4,"venue":"Anthm","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":5,"venue":"Athletic Club & Social","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":6,"venue":"Hennies","channel":"Private Sales","firstName":"","lastName":"","location":"Stellenbosch","distributor":"","email":"","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"Medium","lastContacted":"2026-02-10","status":"New","notes":"Parking for now","totalSales":0},{"id":7,"venue":"Beyond","channel":"Private Sales","firstName":"Sebastian","lastName":"Stehr","location":"Constantia","distributor":"","email":"chef@beyondrestaurant.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-10","status":"New","notes":"Follow up in June","totalSales":0},{"id":8,"venue":"Rust en Vrede","channel":"Private Sales","firstName":"Fabio","lastName":"Daniel","location":"Stellenbosch","distributor":"","email":"fabio@rustenvrede.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-10","status":"Converted","notes":"Rejected","totalSales":0},{"id":9,"venue":"Blue Peter","channel":"Private Sales","firstName":"Marshal","lastName":"","location":"Blouberg","distributor":"","email":"","phone":"083 981 8770","leadSource":"Referral","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"Converted","notes":"","totalSales":0},{"id":10,"venue":"The Tasting Room at Creation","channel":"Private Sales","firstName":"Eleanor","lastName":"Niehaus","location":"Hermanus","distributor":"","email":"eleanor@creationwines.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-10","status":"New","notes":"Rejected","totalSales":0},{"id":11,"venue":"Dusk","channel":"Private Sales","firstName":"Callen","lastName":"Austin","location":"Stellenbosch","distributor":"","email":"","phone":"081 409 6544","leadSource":"Referral","accountManager":"Alex","priority":"Low","lastContacted":"2026-03-11","status":"Contacted","notes":"Rejected","totalSales":0},{"id":12,"venue":"Firemans Arms","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-20","status":"New","notes":"Promo finished. Not a focus.","totalSales":0},{"id":13,"venue":"Chefs Table","channel":"Private Sales","firstName":"Karl","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"2026-04-20","status":"Converted","notes":"Calling today","totalSales":0},{"id":14,"venue":"Bossa Group","channel":"Private Sales","firstName":"Chris","lastName":"","location":"Paarl","distributor":"","email":"","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-03-24","status":"Prospect","notes":"Parking for now","totalSales":0},{"id":15,"venue":"Chefs Warehouse Bree","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":16,"venue":"Aubergine","channel":"Private Sales","firstName":"Harald","lastName":"Bresselschmidt","location":"CBD","distributor":"","email":"chef@aubergine.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"Converted","notes":"Meeting on 24th","totalSales":0},{"id":17,"venue":"Chorus","channel":"Private Sales","firstName":"Mareli","lastName":"Basson","location":"Somerset West","distributor":"","email":"mareli@bertusbasson.com / info@bertusbasson.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Calling today","totalSales":0},{"id":18,"venue":"Club Heritage","channel":"Private Sales","firstName":"Pholoso","lastName":"","location":"","distributor":"","email":"","phone":"073 798 1918","leadSource":"Referral","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-17","status":"Converted","notes":"","totalSales":0},{"id":19,"venue":"Club Soho","channel":"Private Sales","firstName":"Krys","lastName":"","location":"","distributor":"","email":"djkrys86@gmail.com","phone":"083 533 9789","leadSource":"Referral","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"New","notes":"","totalSales":0},{"id":20,"venue":"Foxcroft","channel":"Private Sales","firstName":"Zoe","lastName":"Wepener","location":"","distributor":"","email":"beauadmin@chefswarehouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Emailed. Waiting for response.","totalSales":0},{"id":21,"venue":"Cyra","channel":"Private Sales","firstName":"Candice","lastName":"Philip","location":"Jo'Burg","distributor":"","email":"chefcandicephilip@gmail.com","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-03-10","status":"New","notes":"","totalSales":0},{"id":22,"venue":"Gigi","channel":"Private Sales","firstName":"Moses","lastName":"Moloi","location":"CBD","distributor":"","email":"moses@gigirestaurant.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"No response. Call.","totalSales":0},{"id":23,"venue":"Grand Cafe","channel":"Private Sales","firstName":"Rees","lastName":"Maxwell","location":"Sea Point","distributor":"","email":"","phone":"082 874 2299","leadSource":"Referral","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"Prospect","notes":"Struggles to get hold of Rees. Will keep persisting","totalSales":0},{"id":24,"venue":"Grub & Vne","channel":"Private Sales","firstName":"Matthew","lastName":"Manning","location":"","distributor":"","email":"info@mattmanningchef.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Emailed. Waiting for response.","totalSales":0},{"id":25,"venue":"H\u014dseki","channel":"Private Sales","firstName":"Virgil","lastName":"Kahn","location":"Stellenbosch","distributor":"","email":"lodge.chef@delaire.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"No response. Call.","totalSales":0},{"id":26,"venue":"La Colombe","channel":"Private Sales","firstName":"James","lastName":"Gaag","location":"Hout Bay","distributor":"","email":"james@lacolombe.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":27,"venue":"Post & Pepper","channel":"Private Sales","firstName":"Jess","lastName":"van Dyk","location":"Stellenbosch","distributor":"","email":"jess@postandpepper.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"Converted","notes":"Follow up in June","totalSales":0},{"id":28,"venue":"Serendipity","channel":"Private Sales","firstName":"Lizelle","lastName":"Stolze","location":"Wiilderness","distributor":"","email":"chef123@mweb.co.za/\nrestaurant@serendipitywilderness.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Call.","totalSales":0},{"id":29,"venue":"aha Hotels & Lodges","channel":"Private Sales","firstName":"Ruzandri","lastName":"","location":"","distributor":"","email":"ruzandri.stoltz@aha.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":30,"venue":"Hyatt Regency Cape Town","channel":"Private Sales","firstName":"Sasha","lastName":"Lewis","location":"","distributor":"","email":"sasha.lewis@hyatt.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":31,"venue":"Fac Cactus - Blouberg","channel":"Private Sales","firstName":"Lesley","lastName":"","location":"Blouberg","distributor":"","email":"simon@fatcactus.co.za","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"High","lastContacted":"2026-01-22","status":"New","notes":"","totalSales":0},{"id":32,"venue":"Silo (RP)","channel":"Private Sales","firstName":"Gerrit","lastName":"","location":"","distributor":"","email":"gerrit@trp.travel","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":33,"venue":"Steenberg GC","channel":"Private Sales","firstName":"Tatiana","lastName":"","location":"","distributor":"","email":"fbmanager@steenberggolfclub.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":34,"venue":"Twelve Apostles Hotel","channel":"Private Sales","firstName":"Mo","lastName":"Taliet","location":"","distributor":"","email":"groupproman12a@rchmail.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":35,"venue":"Westlake GC","channel":"Private Sales","firstName":"Tina","lastName":"","location":"","distributor":"","email":"events@westlakegolfclub.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":36,"venue":"De Zalze GC","channel":"Private Sales","firstName":"Ian","lastName":"","location":"","distributor":"","email":"FBManager@dezalzegolf.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":37,"venue":"Pearl Valley GC","channel":"Private Sales","firstName":"Stafford","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":38,"venue":"Arabella GC","channel":"Private Sales","firstName":"Sterlia","lastName":"Van Der Merwe","location":"","distributor":"","email":"sterlia.vandermerwe@southernsun.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":39,"venue":"Red Carnation Hotels","channel":"Private Sales","firstName":"Jameel","lastName":"","location":"","distributor":"","email":"foodbeverage@oysterbox.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":40,"venue":"Royal Johannesburg GC","channel":"Private Sales","firstName":"Erik","lastName":"","location":"","distributor":"","email":"fb@royaljhb.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":41,"venue":"Happy Folks","channel":"Private Sales","firstName":"Jacobus","lastName":"","location":"Stellenbosch / Blouberg","distributor":"","email":"blouberg@happyfolks.co.za","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"Low","lastContacted":"2026-01-20","status":"New","notes":"Need to send him a email and get Franchise approval to list Avante.","totalSales":0},{"id":42,"venue":"Madame Zingara Group","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":43,"venue":"More Family Collection","channel":"Private Sales","firstName":"Lara","lastName":"","location":"","distributor":"","email":"lara@more.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":44,"venue":"The Bay Hotel","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":45,"venue":"Premier Hotels","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":46,"venue":"Kaapstadt Brauhaus","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-21","status":"New","notes":"","totalSales":0},{"id":47,"venue":"Paarl Brandy Bar","channel":"Private Sales","firstName":"Allan","lastName":"","location":"Paarl","distributor":"","email":"","phone":"082 447 2538","leadSource":"Event","accountManager":"Alex","priority":"Low","lastContacted":"2026-04-02","status":"Qualified","notes":"Opening bar in Feb. Need to set up eeting this week","totalSales":0},{"id":48,"venue":"De Tafel","channel":"Private Sales","firstName":"Gregory","lastName":"Henderson","location":"Wynberg","distributor":"","email":"execchef@palmhouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"New","notes":"Waiting on response","totalSales":0},{"id":49,"venue":"Delaire Graff Restaurant","channel":"Private Sales","firstName":"Clinton","lastName":"Jacobs","location":"Stellenbosch","distributor":"","email":"dgr.headchef@delaire.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"New","notes":"No response. Call.","totalSales":0},{"id":50,"venue":"Fable","channel":"Private Sales","firstName":"Josh","lastName":"Sarembock","location":"CBD","distributor":"","email":"","phone":"060 313 4539","leadSource":"Event","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"Prospect","notes":"Keen to stock. Contracts end in June.","totalSales":0},{"id":51,"venue":"Fermier","channel":"Private Sales","firstName":"Adriaan","lastName":"Maree","location":"Pretoria","distributor":"","email":"bookings@fermierrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"New","notes":"No response. Call.","totalSales":0},{"id":52,"venue":"Salsify at the Roundhouse","channel":"Private Sales","firstName":"Ryan","lastName":"Cole","location":"Camps Bay","distributor":"","email":"ryan@salsify.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Figuring out distribution","totalSales":0},{"id":53,"venue":"Les Cr\u00e9atifs","channel":"Private Sales","firstName":"Wandile","lastName":"Mabaso","location":"Jo'Burg","distributor":"","email":"chefwmabaso@gmail.com","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-10","status":"Converted","notes":"Meeting end of May","totalSales":0},{"id":54,"venue":"Terrarium","channel":"Private Sales","firstName":"Anlou","lastName":"Erasmus","location":"Waterfront","distributor":"","email":"qvchef@queenvictoriahotel.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":55,"venue":"The Chefs' Table","channel":"Private Sales","firstName":"Mathew","lastName":"Armbruster","location":"","distributor":"","email":"mathew@thechefstable.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":56,"venue":"The Jordan Restaurant with Marthinus Ferreira","channel":"Private Sales","firstName":"Leigh","lastName":"Simmidari","location":"Stellenbosch","distributor":"","email":"restaurant@jordanwines.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Calling today","totalSales":0},{"id":57,"venue":"Upper Union","channel":"Private Sales","firstName":"Amori","lastName":"Burger","location":"CBD","distributor":"","email":"amori@upperunion.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"Converted","notes":"Followed up. Waiting.","totalSales":0},{"id":58,"venue":"Chefs Warehouse Maison","channel":"Private Sales","firstName":"David Schneider","lastName":"Xavier Francis (head chef)","location":"Franschoek","distributor":"","email":"david@chefswarehouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-06","status":"New","notes":"No response. Call.","totalSales":0},{"id":59,"venue":"Medusa","channel":"Private Sales","firstName":"Vuyani","lastName":"","location":"Milnerton","distributor":"","email":"medusaloungerooftop@gmail.com","phone":"072 258 2463","leadSource":"Cold Call","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-16","status":"Converted","notes":"Done Deals","totalSales":0},{"id":60,"venue":"MIlnerton Golf Club","channel":"Private Sales","firstName":"Alfie","lastName":"Schneeburger","location":"Milnerton","distributor":"","email":"catering@milgolf.co.za","phone":"082 576 7887","leadSource":"Event","accountManager":"Loydz","priority":"Low","lastContacted":"2026-01-20","status":"New","notes":"","totalSales":0},{"id":61,"venue":"Misthom Lounge","channel":"Private Sales","firstName":"Lee","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"Low","lastContacted":"2026-02-20","status":"New","notes":"","totalSales":0},{"id":62,"venue":"Kove Collection","channel":"Private Sales","firstName":"Benito","lastName":"","location":"Camps Bay","distributor":"Karino","email":"","phone":"063 212 7038","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-07","status":"Prospect","notes":"Benito likes Avante. We have a relationship. Waiting for an order and which restaurants.","totalSales":0},{"id":63,"venue":"Mozambik","channel":"Private Sales","firstName":"Uhuru","lastName":"","location":"","distributor":"","email":"","phone":"084 057 0572","leadSource":"Walk into Store","accountManager":"Loydz","priority":"High","lastContacted":"2026-01-21","status":"New","notes":"","totalSales":0},{"id":64,"venue":"Noah","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":65,"venue":"Eike","channel":"Private Sales","firstName":"Michael","lastName":"Fuller","location":"Stellenbosch","distributor":"","email":"kitchen@eikerestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"Callling today","totalSales":0},{"id":66,"venue":"FABER","channel":"Private Sales","firstName":"Dale","lastName":"Stevens","location":"Paarl","distributor":"","email":"faber@avondalewine.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"No response. Call.","totalSales":0},{"id":67,"venue":"TTK Fledgelings","channel":"Private Sales","firstName":"Nathan","lastName":"Clarke","location":"CBD","distributor":"","email":"steven@lukedaleroberts.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":68,"venue":"Life and Brand Portfolo","channel":"Private Sales","firstName":"Jason","lastName":"","location":"Cape Town","distributor":"","email":"jason@lifeandbrand.co.za","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-08","status":"New","notes":"Waiting for response","totalSales":0},{"id":69,"venue":"Spice Route","channel":"Private Sales","firstName":"Gigi","lastName":"Bisogno","location":"Paarl","distributor":"","email":"","phone":"083 441 9870","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-08","status":"Contacted","notes":"","totalSales":0},{"id":70,"venue":"Qunu","channel":"Private Sales","firstName":"Matthew","lastName":"Foxon","location":"Jo'Burg","distributor":"","email":"matthewf@saxon.co.za/scottd@saxon.co.za","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-10","status":"New","notes":"Given to Loydz","totalSales":0},{"id":71,"venue":"Room13, Savage & Concept","channel":"Private Sales","firstName":"Prince","lastName":"","location":"","distributor":"","email":"","phone":"061 982 6349","leadSource":"Referral","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"New","notes":"","totalSales":0},{"id":72,"venue":"La Petite Colombe","channel":"Private Sales","firstName":"Peter","lastName":"Duncan","location":"Franschoek","distributor":"","email":"peter@lapetitecolombe.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-08","status":"Lost","notes":"Revisit in Q2","totalSales":0},{"id":73,"venue":"Le coin Fran\u00e7ais","channel":"Private Sales","firstName":"Darren","lastName":"Badenhorst","location":"Franschoek","distributor":"","email":"darren@lecoinfrancais.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-08","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":74,"venue":"PIER","channel":"Private Sales","firstName":"John","lastName":"Norris-Rogers","location":"Waterfront","distributor":"","email":"john@pier.restaurant","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-08","status":"New","notes":"Callling today","totalSales":0},{"id":75,"venue":"Seat - Blouberg","channel":"Private Sales","firstName":"Anton","lastName":"","location":"Blouberg","distributor":"","email":"info@seatrestaurant.com","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-01-22","status":"New","notes":"","totalSales":0},{"id":76,"venue":"LifeGrande Group","channel":"Private Sales","firstName":"Jeanel / John","lastName":"","location":"","distributor":"","email":"jeanel@lifegrandcafe.com / john@lifegrandcafe.com","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-09","status":"Prospect","notes":"Waiting for response","totalSales":0},{"id":77,"venue":"Stellenbosch GC","channel":"Private Sales","firstName":"Gerhard","lastName":"","location":"Stellenbosch","distributor":"Karino","email":"chef@stbgolf.com","phone":"072 198 7912","leadSource":"Event","accountManager":"Alex","priority":"High","lastContacted":"2026-04-09","status":"Prospect","notes":"Waitng for order coonfirmation","totalSales":0},{"id":78,"venue":"Galjoen","channel":"Private Sales","firstName":"Isca","lastName":"Stoltz","location":"CBD","distributor":"","email":"eat@galjoencpt.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-09","status":"New","notes":"No response. Call.","totalSales":0},{"id":79,"venue":"UCT Bar","channel":"Private Sales","firstName":"Rob","lastName":"Munroe","location":"UCT","distributor":"","email":"","phone":"079 569 4688","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-13","status":"Converted","notes":"Confirmed","totalSales":0},{"id":80,"venue":"SoHo","channel":"Private Sales","firstName":"DJ Krys","lastName":"","location":"Long Street","distributor":"","email":"083 533 9789","phone":"083 533 9789","leadSource":"Referral","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-21","status":"Converted","notes":"","totalSales":0},{"id":81,"venue":"Belly of the Beast","channel":"Private Sales","firstName":"Anouchka","lastName":"Horn","location":"CBD","distributor":"","email":"eat@bellyofthebeast.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-13","status":"New","notes":"No response. Call.","totalSales":0},{"id":82,"venue":"Cavalli","channel":"Private Sales","firstName":"Lucas","lastName":"Carstens","location":"Stellenbosch","distributor":"","email":"chef@cavalliestate.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-13","status":"New","notes":"Calling today","totalSales":0},{"id":83,"venue":"Embarc","channel":"Private Sales","firstName":"Darren","lastName":"O'Donovan","location":"Jo'Burg","distributor":"","email":"info@embarcrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-13","status":"New","notes":"No response. Call.","totalSales":0},{"id":84,"venue":"Thirsty Scarecrow","channel":"Private Sales","firstName":"Ivan","lastName":"Botha","location":"Stellenbosch","distributor":"Direct","email":"kzetler@gmail.com","phone":"084 263 4243","leadSource":"Event","accountManager":"Alex","priority":"Low","lastContacted":"2026-04-14","status":"Converted","notes":"Need to drop barrels. Waiting for order amount.","totalSales":0},{"id":85,"venue":"&Beyond","channel":"Private Sales","firstName":"Nicole Robinson","lastName":"","location":"","distributor":"","email":"nicole.robinson@andbeyond.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-14","status":"New","notes":"","totalSales":0},{"id":86,"venue":"Talking to Strangers","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":87,"venue":"Wolfgat","channel":"Private Sales","firstName":"Kobus","lastName":"van der Merwe","location":"","distributor":"","email":"jjvandermerwe@gmail.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-15","status":"New","notes":"Calling today","totalSales":0},{"id":88,"venue":"Thornybush","channel":"Private Sales","firstName":"Zanele","lastName":"","location":"","distributor":"","email":"","phone":"157931976","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-15","status":"New","notes":"","totalSales":0},{"id":89,"venue":"The Waterside","channel":"Private Sales","firstName":"Roxy","lastName":"Mudie","location":"Waterfront","distributor":"","email":"roxy@thewaterside.restaurant","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-16","status":"New","notes":"Calling today","totalSales":0},{"id":90,"venue":"Clovelly Golf Club","channel":"Private Sales","firstName":"Paul","lastName":"","location":"","distributor":"","email":"gm@clovelly.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-16","status":"New","notes":"","totalSales":0},{"id":91,"venue":"The LivingRoom","channel":"Private Sales","firstName":"Johannes","lastName":"Richter","location":"Durban","distributor":"","email":"info@summerhillkzn.com","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-01-22","status":"New","notes":"Given to Loydz","totalSales":0},{"id":92,"venue":"The Pot Luck Club, Johannesburg","channel":"Private Sales","firstName":"Ebie du Toit/ Vicky Peech","lastName":"Ebie du Toit/ Vicky Peech","location":"Jo'Burg","distributor":"","email":"ebi@plcj.co.za traveller@thepeech.co.za","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-01-22","status":"New","notes":"","totalSales":0},{"id":93,"venue":"Marriott / Protea Hotels","channel":"Private Sales","firstName":"Unathi","lastName":"Dyonase","location":"","distributor":"","email":"Unathi.dyonase@marriot.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-17","status":"New","notes":"","totalSales":0},{"id":94,"venue":"Chefs Warehouse Beau Constantia","channel":"Private Sales","firstName":"Ivor","lastName":"Jones","location":"Constantia","distributor":"","email":"ivor@chefswarehouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-20","status":"New","notes":"Followed up in March","totalSales":0},{"id":95,"venue":"COY","channel":"Private Sales","firstName":"Teenola","lastName":"Govender","location":"CBD","distributor":"","email":"teenola@coyrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-20","status":"New","notes":"Callling today","totalSales":0},{"id":96,"venue":"The Red Room","channel":"Private Sales","firstName":"Caroline","lastName":"Lamb","location":"CBD","distributor":"","email":"caroline@cwredroom.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-21","status":"Converted","notes":"Meeting next week","totalSales":0},{"id":97,"venue":"Tsogo Sun Mossel Bay","channel":"Private Sales","firstName":"George","lastName":"Webber","location":"Mossel Bay","distributor":"","email":"","phone":"838227550","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-21","status":"New","notes":"","totalSales":0},{"id":98,"venue":"\u00eblgr.","channel":"Private Sales","firstName":"Jesper","lastName":"Nillson","location":"CBD","distributor":"","email":"jesper@elgr.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-22","status":"New","notes":"Callling today","totalSales":0},{"id":99,"venue":"FYN","channel":"Private Sales","firstName":"Peter","lastName":"Templehoff","location":"CBD","distributor":"","email":"tempelhoff@gmail.com/ Kerry@fynrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-22","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":100,"venue":"Quay 4","channel":"Private Sales","firstName":"Leroy","lastName":"","location":"Waterfront","distributor":"","email":"","phone":"660793944","leadSource":"Walk into Store","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-22","status":"New","notes":"Waiting response for a meeting with owner","totalSales":0},{"id":101,"venue":"Southern Sun","channel":"Private Sales","firstName":"David","lastName":"Mokwebo","location":"","distributor":"","email":"david.mokwebo@southernsun.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"","totalSales":0},{"id":102,"venue":"Unclaimed","channel":"Private Sales","firstName":"Aidan","lastName":"","location":"Kloof Street","distributor":"","email":"","phone":"079 380 9210","leadSource":"Networking","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"Converted","notes":"","totalSales":0},{"id":103,"venue":"BMW Dealership","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"","totalSales":0},{"id":104,"venue":"Bidvest Lounge","channel":"Private Sales","firstName":"Suzanne","lastName":"VD","location":"","distributor":"","email":"SuzanneVD@bidvestcatering.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"Sent to procurement department","totalSales":0},{"id":105,"venue":"Molenvlliet","channel":"Private Sales","firstName":"Juliene","lastName":"","location":"","distributor":"","email":"stay@molenvliet.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":106,"venue":"Babylonstoren","channel":"Private Sales","firstName":"Elana","lastName":"","location":"","distributor":"","email":"functions@babylonstoren.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"High","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":107,"venue":"Boschendal","channel":"Private Sales","firstName":"Shanel","lastName":"","location":"","distributor":"","email":"functions@babylonstoren.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"High","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":108,"venue":"Da'Aria","channel":"Private Sales","firstName":"Danielle","lastName":"","location":"","distributor":"","email":"events@daria.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"High","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":109,"venue":"Slow Lounge","channel":"Private Sales","firstName":"Janine","lastName":"","location":"","distributor":"","email":"janine@purefoods.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"Waiting for response","totalSales":0},{"id":110,"venue":"Youngblood","channel":"Private Sales","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":111,"venue":"OK Liquors Parow","channel":"Trade Retail","firstName":"Denver","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"844552084","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"New","notes":"Will talk to owner and get back","totalSales":0},{"id":112,"venue":"OK Liquors Tygerdal","channel":"Trade Retail","firstName":"Ray","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"211090099","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":113,"venue":"Tops Bellville","channel":"Trade Retail","firstName":"Yushrah","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"638150626","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":114,"venue":"Tops Boston","channel":"Trade Retail","firstName":"Valencia","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"653915424","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":115,"venue":"Tops Glenwood","channel":"Trade Retail","firstName":"Obert","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"787540057","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":116,"venue":"Tops Parow Valley","channel":"Trade Retail","firstName":"Fiona","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"746320000","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":117,"venue":"Observatory Bottle Store","channel":"Trade Retail","firstName":"Dani","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"719445117","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-16","status":"New","notes":"Not interested yet","totalSales":0},{"id":118,"venue":"Blue Bottle Observatory","channel":"Trade Retail","firstName":"Ling","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-16","status":"New","notes":"Owner will get back to me","totalSales":0},{"id":119,"venue":"Blue Bottle Panorama (Day 3)","channel":"Trade Retail","firstName":"Carl","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"842054604","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":120,"venue":"Market Liquors Parklands Sandowne","channel":"Trade Retail","firstName":"Adriana","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"822325842","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":121,"venue":"Tops Belhar","channel":"Trade Retail","firstName":"Jonothan","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"673236926","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-13","status":"New","notes":"Sent owner a mail, no feedback","totalSales":0},{"id":122,"venue":"Blue Bottle Wild Liquors","channel":"Trade Retail","firstName":"Evan","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"745206017","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":123,"venue":"Liquor City Paarl","channel":"Trade Retail","firstName":"Zheng","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"766767666","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":124,"venue":"Market Liquors Malmesbury","channel":"Trade Retail","firstName":"Janicka","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"840444024","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":125,"venue":"Market Liquors Paarl (Day 1)","channel":"Trade Retail","firstName":"Joy","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"672330172","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":126,"venue":"OK Liquors Rembrandt Mall","channel":"Trade Retail","firstName":"Burton","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"New","notes":"Under new ownership","totalSales":0},{"id":127,"venue":"Tops Zomerlust (Day 1)","channel":"Trade Retail","firstName":"Dimitri","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"769262883","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":128,"venue":"Blue Bottle Liquors Paarl Mall","channel":"Trade Retail","firstName":"Louise","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-02","status":"New","notes":"Not interested yet","totalSales":0},{"id":129,"venue":"Blue Bottle Liquor for Africa","channel":"Trade Retail","firstName":"Yolande","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"836072111","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":130,"venue":"Market Liquors Brackenfell","channel":"Trade Retail","firstName":"Elizabeth","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"612727071","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":131,"venue":"OK Liquors Goedemoed","channel":"Trade Retail","firstName":"Dricus","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"719122379","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":132,"venue":"Tops Brackenfell","channel":"Trade Retail","firstName":"Johan","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"652973469","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":133,"venue":"Tops Kraaifontein (Day 8)","channel":"Trade Retail","firstName":"Adriaan","lastName":"","location":"Cape Town","distributor":"Karino","email":"kraaifonteinsuperspar@retail.co.za","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":134,"venue":"Tops Protea Hoogte (Day 8)","channel":"Trade Retail","firstName":"LJ","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"712375157","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":135,"venue":"Tops Helshoogte","channel":"Trade Retail","firstName":"Patuma","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"763409494","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":136,"venue":"Carolines Tokai","channel":"Trade Retail","firstName":"Caroline","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"New","notes":"Not interested","totalSales":0},{"id":137,"venue":"Market Liquors Hermanus","channel":"Trade Retail","firstName":"JJ","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"798977999","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":138,"venue":"OK Liquors Napier","channel":"Trade Retail","firstName":"Ben","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"822576721","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"New","notes":"No feedback whatsoever","totalSales":0},{"id":139,"venue":"Tops Kleinmond (Day 9)","channel":"Trade Retail","firstName":"Marjana","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"798331149","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":140,"venue":"Tops Villiersdorp (Day 9)","channel":"Trade Retail","firstName":"Elaine","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"662807287","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":141,"venue":"Tops Welgevonden (Day 7)","channel":"Trade Retail","firstName":"Wayne","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"621274008","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":142,"venue":"Market Liquors Tokai (Day 10)","channel":"Trade Retail","firstName":"Kim","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"730617906","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":143,"venue":"OK Liquors Spineroad","channel":"Trade Retail","firstName":"Eleanor","lastName":"","location":"Cape Town","distributor":"Karino","email":"spineroad@ok.co.za","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"New","notes":"Under new ownership","totalSales":0},{"id":144,"venue":"Tops Fishhoek","channel":"Trade Retail","firstName":"Gilbert","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"768108940","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"Listed VS & VSOP","totalSales":0},{"id":145,"venue":"Tops Glencairn","channel":"Trade Retail","firstName":"Theresa","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Cold Call","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":146,"venue":"Tops Houtbay","channel":"Trade Retail","firstName":"Farai","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"710779622","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":147,"venue":"Tops Muizenberg","channel":"Trade Retail","firstName":"Denise","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"820525529","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":148,"venue":"Tops Oakhurst","channel":"Trade Retail","firstName":"Chicco","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"609232264","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":149,"venue":"Tops Weltevreden (Day 10)","channel":"Trade Retail","firstName":"Paul","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"610052357","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":150,"venue":"Blue Bottle Elands Bay","channel":"Trade Retail","firstName":"Patrick","lastName":"","location":"West-Coast","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-23","status":"New","notes":"Will give feedback in April","totalSales":0},{"id":151,"venue":"Tops Edgemead","channel":"Trade Retail","firstName":"Nolene","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"724245680","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-05","status":"New","notes":"Not interested yet","totalSales":0},{"id":152,"venue":"Tops Bredasdorp","channel":"Trade Retail","firstName":"Anna","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"728264575","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-03","status":"Converted","notes":"","totalSales":0},{"id":153,"venue":"Tops Aurora (Day 6)","channel":"Trade Retail","firstName":"Alcino","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"765381927","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":154,"venue":"Tops Durbanville","channel":"Trade Retail","firstName":"Devon","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"815151023","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":155,"venue":"Tops Haasendal","channel":"Trade Retail","firstName":"Tyrese","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"677984010","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":156,"venue":"Tops Old Oak","channel":"Trade Retail","firstName":"Hennie","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"219192711","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":157,"venue":"Tops Palm Grove","channel":"Trade Retail","firstName":"Charl","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"219762148","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":158,"venue":"Tops Phesantekraal","channel":"Trade Retail","firstName":"Jestin","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"605484307","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":159,"venue":"Tops Kuilsriver (Day 5)","channel":"Trade Retail","firstName":"Kevin","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"785021144","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":160,"venue":"Tops Sunrise Circle (Day 5)","channel":"Trade Retail","firstName":"Barry","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"738878330","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":161,"venue":"Tops Mowbray","channel":"Trade Retail","firstName":"Michelle","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"765520284","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":162,"venue":"Tops Observatory","channel":"Trade Retail","firstName":"Lincon","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"658430389","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":163,"venue":"Tops Wex1","channel":"Trade Retail","firstName":"Ernst","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"738837072","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":164,"venue":"Tops Woodstock Quarter","channel":"Trade Retail","firstName":"Ernst","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"738837072","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":165,"venue":"Whisky Shop","channel":"Trade Retail","firstName":"Hector","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"833771113","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":166,"venue":"Liquor City Long Street (Day 4)","channel":"Trade Retail","firstName":"Arthur","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"680530467","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":167,"venue":"Woodstock Liquors","channel":"Trade Retail","firstName":"Inacio","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"832689430","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":168,"venue":"Tops Adderly","channel":"Trade Retail","firstName":"Noni","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":169,"venue":"Tops Cape Station","channel":"Trade Retail","firstName":"Noni","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":170,"venue":"Tops Century Village","channel":"Trade Retail","firstName":"Anthony","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"825215452","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":171,"venue":"Blouberg Liquors","channel":"Trade Retail","firstName":"Inacio","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"832689430","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":172,"venue":"NGF Sunset Beach","channel":"Trade Retail","firstName":"Gary","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"726145099","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":173,"venue":"OK Liquors Duynefontein","channel":"Trade Retail","firstName":"Frean","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"722437068","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":174,"venue":"Tops Blouberg","channel":"Trade Retail","firstName":"Canais","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"635074577","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":175,"venue":"Tops Melkbos","channel":"Trade Retail","firstName":"Abigail","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"715768900","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":176,"venue":"Tops Milnerton","channel":"Trade Retail","firstName":"Cheslyn","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"692649715","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":177,"venue":"Tops Morningfield (Day 3)","channel":"Trade Retail","firstName":"Sinetemba","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"833584184","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":178,"venue":"Tops Parklands","channel":"Trade Retail","firstName":"Morne","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"743749296","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":179,"venue":"Tops Riberios","channel":"Trade Retail","firstName":"Louise","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"673611248","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":180,"venue":"Tops Royal Ascot","channel":"Trade Retail","firstName":"Tumeka","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"632692784","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":181,"venue":"Tops Sunningdale","channel":"Trade Retail","firstName":"Tinotenda","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"658334199","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":182,"venue":"Tops Tableview","channel":"Trade Retail","firstName":"Roger","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"718921187","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":183,"venue":"Tops Malmesbury","channel":"Trade Retail","firstName":"Hendika","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"785751687","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":184,"venue":"Tops Vineyard","channel":"Trade Retail","firstName":"Christiaan","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"720407583","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":185,"venue":"Tops Cape Quarter (Day 2)","channel":"Trade Retail","firstName":"Chantal","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"836613653","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":186,"venue":"Barkeeper","channel":"Trade Retail","firstName":"Carla","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"823391079","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":187,"venue":"Constantia Wine & Craft","channel":"Trade Retail","firstName":"Jeremy","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"786424044","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":188,"venue":"NGF Gardens","channel":"Trade Retail","firstName":"Vuyo","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"7433425032","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":189,"venue":"NGF Three Anchor Bay","channel":"Trade Retail","firstName":"Joline","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"786584312","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":190,"venue":"Tops Bergvliet","channel":"Trade Retail","firstName":"Vuyi","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"766611750","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"Not interested yet","totalSales":0},{"id":191,"venue":"Tops Kloof","channel":"Trade Retail","firstName":"Makhumalo","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"631053550","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":192,"venue":"Tops Rosmead (Day 2)","channel":"Trade Retail","firstName":"Rochelle","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"728603028","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":193,"venue":"Tops Seapoint","channel":"Trade Retail","firstName":"Stuart","lastName":"","location":"CBD","distributor":"Karino","email":"seapoint1@retail.spar.co.za","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":194,"venue":"Tops Vredehoek","channel":"Trade Retail","firstName":"Lisa","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"682316301","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":195,"venue":"Tops Bottelary (Day 7)","channel":"Trade Retail","firstName":"Jestin","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"605484307","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-27","status":"Converted","notes":"","totalSales":0},{"id":196,"venue":"OK Liquors Urban Sonstraal","channel":"Trade Retail","firstName":"Bri-Niel","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"733719691","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":197,"venue":"Tops Cape Gate","channel":"Trade Retail","firstName":"Frank","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"638538885","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":198,"venue":"Liquor City Somerset Mall","channel":"Trade Retail","firstName":"Charlene","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"715243228","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":199,"venue":"OK Liquors Strand","channel":"Trade Retail","firstName":"Meaghan","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"733695375","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":200,"venue":"Tops Cinnamon","channel":"Trade Retail","firstName":"Tafadswa","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"614817661","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":201,"venue":"Tops De Jonker","channel":"Trade Retail","firstName":"Charlton","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"641629569","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":202,"venue":"Tops Die Boord","channel":"Trade Retail","firstName":"Gershwin","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"613656818","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":203,"venue":"Tops Helderbosch","channel":"Trade Retail","firstName":"Chantal","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"674027744","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":204,"venue":"Tops Helderveu","channel":"Trade Retail","firstName":"Heinrich","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"619525640","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":205,"venue":"Tops Lions Square","channel":"Trade Retail","firstName":"Barbara","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"790973527","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":206,"venue":"Tops Mountainview","channel":"Trade Retail","firstName":"Melisia","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"740583958","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":207,"venue":"Tops Paradyskloof","channel":"Trade Retail","firstName":"Kenny","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"739506767","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":208,"venue":"Tops Paul Roos","channel":"Trade Retail","firstName":"Nathely","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"780743365","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":209,"venue":"Tops Strand","channel":"Trade Retail","firstName":"Lycken","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"620494468","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"New","notes":"Not interesteed yet","totalSales":0},{"id":210,"venue":"Tops Twin Palms","channel":"Trade Retail","firstName":"Senobia","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"843573643","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":211,"venue":"Liquor City Claremont","channel":"Trade Retail","firstName":"Derek","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"216741478","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":212,"venue":"OK Liquors Ridgeworth","channel":"Trade Retail","firstName":"Quinton","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"837799296","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":213,"venue":"Picardi Rebel Claremont (Day 4)","channel":"Trade Retail","firstName":"Tristian","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"662086625","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":214,"venue":"Tops Eastcliff","channel":"Trade Retail","firstName":"Nadia","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"790927628","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":215,"venue":"Tops Eersteriver (Day 6)","channel":"Trade Retail","firstName":"Sinalo","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"735920264","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"New","notes":"Not interested at all","totalSales":0},{"id":216,"venue":"OK Liquors Strandfontein","channel":"Trade Retail","firstName":"Jenique","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"726130092","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":217,"venue":"Tops Alphen","channel":"Trade Retail","firstName":"Riyaaz","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"742624147","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":218,"venue":"Ultra Nyanga","channel":"Trade Retail","firstName":"Sisa","lastName":"Liwani","location":"Nyanga","distributor":"","email":"nyanga@ultraliquors.co.za","phone":"082 398 0404","leadSource":"Online","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-18","status":"New","notes":"","totalSales":0},{"id":219,"venue":"Tops Sonstraal","channel":"Trade Retail","firstName":"Selvinia","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"634486222","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":220,"venue":"Tops Vredekloof","channel":"Trade Retail","firstName":"Riaan","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"605905596","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":221,"venue":"Tops Welgelegeen","channel":"Trade Retail","firstName":"Yianakis","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"720615896","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":222,"venue":"Tops Zevenwacht","channel":"Trade Retail","firstName":"Wesley","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"824777793","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":223,"venue":"Dal Italia Restaurant","channel":"Trade Retail","firstName":"Fabio","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"823372339","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-06","status":"Converted","notes":"","totalSales":0},{"id":224,"venue":"City Lodge Hotels","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":225,"venue":"Anew Hotels","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":226,"venue":"Bon Hotels","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":227,"venue":"Steyn City","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":228,"venue":"Taj Hotel","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":229,"venue":"Glendower","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":230,"venue":"Randpark Golf Club","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":231,"venue":"Houghton Golf Club","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":232,"venue":"Londolozi","channel":"Private Sales","firstName":"Duncan","lastName":"","location":"","distributor":"","email":"duncan@londolozi.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"No response, call today","totalSales":0},{"id":233,"venue":"MalaMala","channel":"Private Sales","firstName":"Rufie or Ross","lastName":"","location":"","distributor":"","email":"","phone":"781008810","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"No response, call today","totalSales":0},{"id":234,"venue":"Shamwari Group","channel":"Private Sales","firstName":"Janine","lastName":"","location":"","distributor":"","email":"procurement@shamwari.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"","totalSales":0},{"id":235,"venue":"Singita (Group)","channel":"Private Sales","firstName":"Schwess","lastName":"","location":"","distributor":"","email":"capepremierwine@singita.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"","totalSales":0},{"id":236,"venue":"Foodbarn Group","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":237,"venue":"Hussar Grill Group","channel":"Private Sales","firstName":"Bradley","lastName":"","location":"","distributor":"","email":"bradleyh@spurcorp.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":238,"venue":"Newmark Hotels","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":239,"venue":"Burkenhead House (RP)","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":240,"venue":"Ellerman House","channel":"Private Sales","firstName":"Manuel","lastName":"","location":"","distributor":"","email":"manuel@ellerman.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":241,"venue":"Fairmont / Accor","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":242,"venue":"Hilton SA","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":243,"venue":"Kapama Private Game Reserve","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":244,"venue":"La Residence (RP)","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":245,"venue":"One&Only Cape Town","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":246,"venue":"Royal Malewane (RP)","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":247,"venue":"Sun International","channel":"Private Sales","firstName":"Mpho","lastName":"Moshotle","location":"","distributor":"","email":"Mpho.moshotle@suninternational.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":248,"venue":"Ulusaba (Virgin)","channel":"Private Sales","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":249,"venue":"Energy Masterbuilders Paarl","channel":"Trade Retail","firstName":"Johan","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":250,"venue":"Hoekstra Farming","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":251,"venue":"Nexus Agriculture","channel":"Trade Retail","firstName":"Ester","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":252,"venue":"PSG Paarl","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":253,"venue":"PSG Stellenbosch","channel":"Trade Retail","firstName":"Joaan","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":254,"venue":"BMW Paarlberg","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":255,"venue":"BMW Stellenbosch","channel":"Trade Retail","firstName":"Christopher","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":256,"venue":"GWM Paarl","channel":"Trade Retail","firstName":"Brandon","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":257,"venue":"Exsa","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":258,"venue":"Paarl Boys High Waterpolo","channel":"Trade Retail","firstName":"Ross","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":259,"venue":"Stellenberg High School","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":260,"venue":"New Orleans High School","channel":"Trade Retail","firstName":"Carlene","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":261,"venue":"Nuweland Paarl","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":262,"venue":"Giflo Group","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":263,"venue":"Samsung","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":264,"venue":"Icon Fruit","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":265,"venue":"Isipani Paarl","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":266,"venue":"Sky vines","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":267,"venue":"Akura Paarl","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":268,"venue":"Rawson H/Q","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":269,"venue":"Workshop 17","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":270,"venue":"Cargokwik","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":271,"venue":"Zuus Paarl","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":272,"venue":"Selavie (Val de Vie)","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":273,"venue":"Blacksmith Kitchen","channel":"Trade Retail","firstName":"Bianca","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":274,"venue":"Grande Roche Hotel","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":275,"venue":"Noop","channel":"Trade Retail","firstName":"Zian","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":276,"venue":"Laborie Estate","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":277,"venue":"Veld and Vine","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":278,"venue":"Val de Vie","channel":"Trade Retail","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0}];
const LOGO_URL = "data:image/jpeg;base64,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ZvrF0Hqj3S48O4jG5jpltJe5N0RFHkmEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAUA1b1bxrSHB33y+SeNVy7x0FtjcBLVyAdB6NHIuf0A9SQC1b1bxrSHB33y+SeNVy7soLbG4CWrkA6D0aORc88gPUkA1l6g6g5LqZnFTlGUVnjVUvkihZuIqaMHyxRt7NH2k7k7kqRsbF13ql3fuRXEeIq2WiG8n7DUHUHJdTM4qcoyis8apl8sULNxFTRg+WKNvZo+0ncnclRZEVmjFRWlcioTm5tyk8thERemIREQBERAEREAREQBERAEREAREQBERAFMNNdSsm0rzqDJsZqeF42ZU0khPg1cW+5jkA7eh6tPMKHosZRUk4yWxlCcoSUovDRbLpbqljOrOCxZFjs/C8bR1lDI4eNRy7bljx+sO6OHMdwJuqk9NdSsm0rzqDJsZqeF42ZU0khPg1cW+5jkA7eh6tPMKzXS3VLGdWcFiyLHZ+F42jrKGRw8ajl23LHj9Yd0cOY7gVm+sXQeqPdLhw7iMbmOmW0l7k3REUeSgREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAFANW9W8a0hwd98vknjVcu7KC2xuAlq5AOg9GjkXP6AepIBat6t41pDg775fJPGq5d46C2xuAlq5AOg9GjkXPPID1JANZeoOoOS6mZxU5RlFZ41TL5YoWbiKmjB8sUbezR9pO5O5KkbGxdd6pd37kVxHiKtlohvJ+w1B1ByXUzOKnKMorPGqZfLFCzcRU0YPlijb2aPtJ3J3JUWRFZoxUVhcioTm5tyk8thERemIREQBERAEREAREQBERAEREAREQBERAEREAREQBTDTXUrJtK86gybGanheNmVNJIT4NXFvuY5AO3oerTzCh6LGUVJaZcjKE5QkpReGi2XS3VLGdWcFiyLHZ+F42jrKGRw8ajl23LHj9Yd0cOY7gTdVJ6a6lZNpZnMGTYzU8L27MqaSQnwauLfcxyAdvQ9WnmFZrpbqljOrOCxZFjs/C8bR1lDI4eNRy7bljx+sO6OHMdwKzfWLoPVHulw4dxGNzHTLaS9yboiKPJQIiIA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";

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
const AvanteLogo = ({ className = '', size = 56 }) => (
  <img
    src={LOGO_URL}
    alt="Avante Cape Brandy"
    className={className}
    style={{ width: size, height: size, objectFit: 'contain', display: 'block' }}
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
    const { data: inserted, error } = await supabase
      .from('clients').insert(clientToDb({ id: undefined, ...newClientData }))
      .select().single();
    if (error) { console.error('[addClient]', error); return null; }
    const newClient = clientFromDb(inserted);
    setClients((prev) => [...prev, newClient]);
    return newClient;
  };

  const resetData = () => {
    askConfirm({
      title: 'Reset all data?',
      message: 'This wipes all logged visits, restores the original 278 client list, and resets all targets to defaults.\n\nThis cannot be undone.',
      confirmLabel: 'RESET EVERYTHING',
      danger: true,
      onConfirm: async () => {
        // Clear visits and clients from DB, re-seed
        await supabase.from('visits').delete().neq('id', 0);
        await supabase.from('clients').delete().neq('id', 0);
        const seedRows = SEED_CLIENTS.map(clientToDb);
        for (let i = 0; i < seedRows.length; i += 50) {
          await supabase.from('clients').insert(seedRows.slice(i, i + 50));
        }
        const seedTargets = SALES_REPS.map(rep => targetToDb(rep, DEFAULT_TARGETS[rep]));
        await supabase.from('targets').upsert(seedTargets, { onConflict: 'rep' });
        setClients(SEED_CLIENTS);
        setVisits([]);
        setTargets(DEFAULT_TARGETS);
        showToast('All data reset to defaults', 'success');
      },
    });
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
            <AvanteLogo size={140} />
          </div>
          <p className="mt-6 text-sm tracking-[0.4em]" style={{ color: '#FDB940', fontFamily: "'Cinzel', serif", fontWeight: 600 }}>LOADING CRM</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen" style={{ background: '#FFFEF2', fontFamily: "'Libre Baskerville', Georgia, serif", color: '#003553' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Libre+Baskerville:ital,wght@0,400;0,700;1,400&family=Cinzel:wght@500;600;700;900&display=swap');
        .font-display { font-family: 'Cinzel', 'Copperplate', serif; letter-spacing: 0.08em; }
        .font-body { font-family: 'Libre Baskerville', Georgia, serif; }
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
        .scrollbar-thin::-webkit-scrollbar { width: 4px; height: 4px; }
        .scrollbar-thin::-webkit-scrollbar-track { background: rgba(0,53,83,0.05); }
        .scrollbar-thin::-webkit-scrollbar-thumb { background: rgba(0,53,83,0.3); border-radius: 3px; }
        @keyframes fadeUp { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
        .fade-up { animation: fadeUp 0.4s ease-out forwards; }
        /* Mobile-friendly touch targets */
        @media (max-width: 767px) {
          select, input[type="date"], input[type="text"], input[type="email"],
          input[type="tel"], input[type="number"], textarea {
            font-size: 16px !important; /* prevents iOS zoom on focus */
            min-height: 44px;
          }
          button { min-height: 44px; }
          .mobile-btn { min-height: 48px; }
        }
        /* Smooth modal transitions */
        .modal-slide-up { animation: slideUp 0.25s ease-out forwards; }
        @keyframes slideUp { from { transform: translateY(100%); opacity: 0.8; } to { transform: translateY(0); opacity: 1; } }
        /* Hide scrollbar on horizontal scroll areas */
        .no-scrollbar::-webkit-scrollbar { display: none; }
        .no-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }
      `}</style>

      <Header view={view} setView={setView} onLog={() => setShowLogModal(true)} onReset={resetData} />

      <main className="max-w-[1400px] mx-auto px-4 md:px-6 lg:px-10 py-5 md:py-8 pb-24 md:pb-8">
        {view === 'dashboard' && (
          <Dashboard
            clients={clients}
            visits={monthVisits}
            allVisits={visits}
            targets={targets}
            activeRep={activeRep}
            setActiveRep={setActiveRep}
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
            if (newClientCtx.onCreated) newClientCtx.onCreated(created);
            setNewClientCtx(null);
            showToast(`New client added: ${created.venue} → ${created.accountManager}'s book`);
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
    <div className="fixed inset-0 z-[90] flex items-center justify-center p-4 bg-ink/80 backdrop-blur-sm" onClick={onCancel}>
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
            className="px-5 py-2.5 font-display text-xs tracking-[0.25em] ink hover:bg-ink/5"
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
function Header({ view, setView, onLog, onReset }) {
  const tabs = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'leads', label: 'Clients', icon: Users },
    { id: 'visits', label: 'Visits', icon: ClipboardList },
    { id: 'manager', label: 'Manager', icon: Settings },
  ];
  return (
    <>
      {/* ── Desktop top header ── */}
      <header className="hidden md:block border-b-2 border-ink relative overflow-hidden" style={{ background: '#003553' }}>
        <RaysBackdrop opacity={0.08} />
        <div className="max-w-[1400px] mx-auto px-6 lg:px-10 py-4 relative">
          <div className="flex items-center justify-between gap-4">
            <div className="flex items-center gap-4">
              <AvanteLogo size={52} />
              <div className="border-l border-white/20 pl-4">
                <p className="font-display text-[9px] tracking-[0.4em]" style={{ color: '#FDB940', fontWeight: 600 }}>SALES PERFORMANCE CRM</p>
                <p className="font-display text-[9px] tracking-[0.3em] mt-0.5" style={{ color: '#FFFEF2', opacity: 0.6 }}>DARE TO FORWARD</p>
              </div>
            </div>
            <nav className="flex items-center gap-1 bg-black/20 p-1">
              {tabs.map(t => {
                const Icon = t.icon;
                const active = view === t.id;
                return (
                  <button key={t.id} onClick={() => setView(t.id)}
                    className={`flex items-center gap-2 px-3 py-2 text-[10px] font-display tracking-[0.15em] transition-all ${active ? 'bg-gold ink' : 'text-white/70 hover:text-white'}`}
                    style={{ fontWeight: 600 }}>
                    <Icon className="w-3.5 h-3.5" />
                    {t.id === 'leads' ? 'LEADS & CLIENTS' : t.id === 'visits' ? 'VISIT LOG' : t.id === 'manager' ? 'MANAGER' : 'DASHBOARD'}
                  </button>
                );
              })}
            </nav>
            <div className="flex items-center gap-2">
              <button onClick={onReset} title="Reset" className="text-white/30 hover:text-white/70 text-[9px] tracking-[0.2em] font-display px-2 py-1">RESET</button>
              <button onClick={onLog} className="flex items-center gap-2 bg-copper hover:bg-gold px-4 py-2.5 font-display text-[10px] tracking-[0.25em] transition-colors" style={{ color: '#FFFEF2', fontWeight: 700 }}>
                <Plus className="w-4 h-4" /> LOG VISIT
              </button>
            </div>
          </div>
        </div>
      </header>

      {/* ── Mobile top bar ── */}
      <header className="md:hidden sticky top-0 z-40 border-b-2 border-ink relative overflow-hidden" style={{ background: '#003553' }}>
        <div className="flex items-center justify-between px-4 py-3">
          <div className="flex items-center gap-3">
            <AvanteLogo size={40} />
            <div>
              <p className="font-display text-[8px] tracking-[0.3em]" style={{ color: '#FDB940', fontWeight: 600 }}>AVANTE CRM</p>
              <p className="font-display text-[7px] tracking-[0.2em]" style={{ color: '#FFFEF2', opacity: 0.5 }}>DARE TO FORWARD</p>
            </div>
          </div>
          <button onClick={onLog} className="flex items-center gap-1.5 bg-copper px-3 py-2 font-display text-[10px] tracking-[0.2em]" style={{ color: '#FFFEF2', fontWeight: 700 }}>
            <Plus className="w-4 h-4" /> LOG VISIT
          </button>
        </div>
      </header>

      {/* ── Mobile bottom nav ── */}
      <nav className="md:hidden fixed bottom-0 left-0 right-0 z-40 border-t-2 border-ink flex" style={{ background: '#003553' }}>
        {tabs.map(t => {
          const Icon = t.icon;
          const active = view === t.id;
          return (
            <button key={t.id} onClick={() => setView(t.id)}
              className={`flex-1 flex flex-col items-center justify-center py-2.5 gap-1 transition-all ${active ? '' : 'opacity-50'}`}
              style={{ color: active ? '#FDB940' : '#FFFEF2' }}>
              <Icon className="w-5 h-5" />
              <span className="font-display text-[8px] tracking-[0.15em]" style={{ fontWeight: active ? 700 : 500 }}>{t.label.toUpperCase()}</span>
            </button>
          );
        })}
      </nav>
    </>
  );
}

// =================== Dashboard ===================
function Dashboard({ clients, visits, allVisits, targets, activeRep, setActiveRep }) {
  const repFilter = (rec, key = 'accountManager') => activeRep === 'All' || rec[key] === activeRep;
  const filteredClients = clients.filter(c => repFilter(c));
  const filteredVisits = visits.filter(v => repFilter(v, 'salesRep'));

  const totalSales = filteredClients.reduce((s, c) => s + (c.totalSales || 0), 0);
  const monthSales = filteredVisits.reduce((s, v) => s + (v.saleAmount || 0), 0);
  const visitCount = filteredVisits.length;
  const sold = filteredVisits.filter(v => v.outcome === 'Sold In').length;
  const totalNew = filteredVisits.length;
  const conversionRate = totalNew > 0 ? Math.round((sold / totalNew) * 100) : 0;

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

  const channelVisits = useMemo(() => {
    const out = { 'Private Sales': 0, 'Trade Retail': 0 };
    filteredVisits.forEach(v => { if (out[v.channel] !== undefined) out[v.channel] += 1; });
    return out;
  }, [filteredVisits]);

  const pct = (a, b) => b > 0 ? Math.min(100, Math.round((a / b) * 100)) : 0;

  return (
    <div className="space-y-5 md:space-y-8 fade-up">
      {/* Header — compact on mobile */}
      <div className="pb-4 border-b border-ink/15">
        <p className="font-display text-[10px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>PERFORMANCE OVERVIEW</p>
        <div className="flex items-start justify-between gap-3 mt-1 flex-wrap">
          <h1 className="font-display text-2xl md:text-4xl ink" style={{ fontWeight: 700 }}>
            {new Date().toLocaleDateString('en-ZA', { month: 'long', year: 'numeric' }).toUpperCase()}
          </h1>
          <RepToggle active={activeRep} onChange={setActiveRep} />
        </div>
      </div>

      {/* KPI cards — 2 col on mobile, 4 on desktop */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <KPICard label="Month Sales" value={ZAR(monthSales)} target={activeTargets.revenue} targetValue={ZAR(activeTargets.revenue)} progress={pct(monthSales, activeTargets.revenue)} icon={DollarSign} accent="copper" />
        <KPICard label="Visits" value={visitCount} target={activeTargets.visits} targetValue={`${activeTargets.visits} visits`} progress={pct(visitCount, activeTargets.visits)} icon={Activity} accent="ocean" />
        <KPICard label="Conversion" value={`${conversionRate}%`} subtitle={`${sold} of ${totalNew} sold`} progress={conversionRate} icon={TrendingUp} accent="gold" />
        <KPICard label="Total TD" value={ZAR(totalSales)} subtitle="All-time" icon={Award} accent="ink" />
      </div>

      {/* Channel cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
        <ChannelCard title="Private Sales" subtitle="On-Consumption" icon={Wine} visits={channelVisits['Private Sales']} target={activeTargets.privateSales} />
        <ChannelCard title="Trade Retail" subtitle="Off-Consumption" icon={Briefcase} visits={channelVisits['Trade Retail']} target={activeTargets.tradeRetail} />
      </div>

      {/* Leaderboard + Recent visits stacked on mobile */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <div className="lg:col-span-2 premium-card p-4 md:p-6">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-2">
              <div className="w-2 h-2 bg-copper diamond-clip"></div>
              <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>LEADERBOARD</h2>
            </div>
            <span className="text-[9px] tracking-[0.2em] font-display ocean">THIS MONTH</span>
          </div>
          <Leaderboard clients={clients} visits={allVisits} targets={targets} />
        </div>
        <div className="premium-card p-4 md:p-6">
          <div className="flex items-center gap-2 mb-4">
            <div className="w-2 h-2 bg-copper diamond-clip"></div>
            <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>RECENT VISITS</h2>
          </div>
          <RecentVisits visits={filteredVisits.slice(-5).reverse()} clients={clients} />
        </div>
      </div>

      {/* Pipeline */}
      <div className="premium-card p-4 md:p-6">
        <div className="flex items-center gap-2 mb-4">
          <div className="w-2 h-2 bg-copper diamond-clip"></div>
          <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>PIPELINE {activeRep !== 'All' && `— ${activeRep.toUpperCase()}`}</h2>
        </div>
        <PipelineGrid clients={filteredClients} />
      </div>
    </div>
  );
}

function RepToggle({ active, onChange }) {
  const opts = ['All', ...SALES_REPS];
  return (
    <div className="inline-flex border border-ink/20 bg-cream">
      {opts.map((r, i) => {
        const isActive = active === r;
        return (
          <button
            key={r}
            onClick={() => onChange(r)}
            className={`px-2.5 md:px-4 py-1.5 md:py-2 font-display text-[9px] md:text-[11px] tracking-[0.2em] transition-all ${isActive ? 'bg-ink' : 'hover:bg-ink/5'} ${i > 0 ? 'border-l border-ink/20' : ''}`}
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
          <div className="h-1 bg-ink/10 relative overflow-hidden">
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
        <div className="w-12 h-12 flex-shrink-0 flex items-center justify-center bg-ink/5 border border-ink/10">
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
          <div className="h-1.5 bg-ink/10 relative overflow-hidden">
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
        <div key={s.rep} className="border border-ink/10 p-3 hover:border-copper/40 transition-colors">
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
              <div className="h-1 bg-ink/10"><div className="h-full bg-copper" style={{ width: `${s.pct}%` }}></div></div>
            </div>
            <div>
              <div className="flex justify-between text-[9px] font-display tracking-[0.15em] mb-1">
                <span className="ocean" style={{ fontWeight: 600 }}>VISITS</span>
                <span className="copper" style={{ fontWeight: 700 }}>{s.visitPct}%</span>
              </div>
              <div className="h-1 bg-ink/10"><div className="h-full" style={{ width: `${s.visitPct}%`, background: '#006C90' }}></div></div>
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
          <div className="h-1 bg-ink/10">
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
            <AvanteLogo size={44} />
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
        <div className="mt-3 pt-3 border-t border-ink/10 grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
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
              <div key={rep} className="border border-ink/10 p-3">
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
                  <tr key={rep} className="border-b border-ink/10">
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
              <button onClick={handleRevert} className="px-2 md:px-3 py-2 font-display text-[9px] tracking-[0.15em] text-white/70 hover:text-white" style={{ fontWeight: 600 }}>REVERT</button>
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
      <div className="flex items-center justify-between pb-3 mb-3 border-b border-ink/10">
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
                  className="flex-1 px-2 py-1.5 border border-ink/20 bg-cream font-display text-sm focus:outline-none focus:border-copper ink"
                  style={{ fontWeight: 700 }}
                />
              </div>
              <div className="flex items-center justify-between mt-1.5 text-[10px]">
                <span className="ocean italic">{f.hint} · actual: <span className="ink font-display" style={{ fontWeight: 700 }}>{f.isCurrency ? ZAR(actual) : actual}</span></span>
                <span className="font-display tracking-wider copper flex-shrink-0" style={{ fontWeight: 700 }}>{pct}%</span>
              </div>
              <div className="h-1 bg-ink/10 mt-1">
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
  const [filterStatus, setFilterStatus] = useState('All');

  const filtered = useMemo(() => {
    return clients.filter(c => {
      if (filterRep !== 'All' && c.accountManager !== filterRep) return false;
      if (filterChannel !== 'All' && c.channel !== filterChannel) return false;
      if (filterStatus !== 'All' && c.status !== filterStatus) return false;
      if (search) {
        const q = search.toLowerCase();
        const hay = `${c.venue} ${c.firstName} ${c.lastName} ${c.location} ${c.email} ${c.phone}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [clients, filterRep, filterChannel, filterStatus, search]);

  return (
    <div className="space-y-4 fade-up">
      {/* Page header */}
      <div className="pb-3 border-b border-ink/15">
        <p className="font-display text-[9px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>CLIENT DATABASE</p>
        <h1 className="font-display text-2xl md:text-4xl mt-1 ink" style={{ fontWeight: 700 }}>LEADS &amp; CLIENTS</h1>
        <p className="italic text-xs md:text-sm ocean mt-0.5">{clients.length} venues tracked</p>
      </div>

      {/* Filters */}
      <div className="space-y-2">
        <div className="relative">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 ocean" />
          <input type="text" placeholder="Search venue, contact, location..." value={search} onChange={(e) => setSearch(e.target.value)} className="w-full pl-10 pr-3 py-3 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper" />
        </div>
        <div className="grid grid-cols-2 gap-2">
          <FilterSelect label="Rep" value={filterRep} onChange={setFilterRep} options={['All', ...SALES_REPS, 'Unassigned']} />
          <FilterSelect label="Channel" value={filterChannel} onChange={setFilterChannel} options={['All', ...CHANNELS]} />
        </div>
        {/* Status pills — horizontal scroll on mobile */}
        <div className="flex items-center gap-1.5 overflow-x-auto pb-1 scrollbar-thin">
          {['All', ...STATUSES].map(s => (
            <button key={s} onClick={() => setFilterStatus(s)}
              className={`flex-shrink-0 px-2.5 py-1 text-[10px] font-display tracking-wider transition-colors ${filterStatus === s ? 'bg-ink' : 'border border-ink/20'}`}
              style={{ color: filterStatus === s ? '#FFFEF2' : '#003553', fontWeight: 600 }}>
              {s.toUpperCase()}
            </button>
          ))}
        </div>
        <p className="text-[10px] italic ocean">{filtered.length} of {clients.length} clients</p>
      </div>

      {/* Desktop table — hidden on mobile */}
      <div className="hidden md:block premium-card overflow-hidden">
        <div className="overflow-x-auto scrollbar-thin max-h-[640px] overflow-y-auto">
          <table className="w-full text-sm">
            <thead className="sticky top-0 z-10" style={{ background: '#003553' }}>
              <tr>
                {['Venue', 'Channel', 'Contact', 'Location', 'Rep', 'Status', 'Last Visit', 'Sales TD', ''].map(h => (
                  <th key={h} className="px-3 py-3 text-left font-display text-[10px] tracking-[0.2em]" style={{ color: '#FDB940', fontWeight: 600 }}>{h.toUpperCase()}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(c => (
                <tr key={c.id} className="border-t border-ink/10 hover:bg-gold/10 cursor-pointer transition-colors" onClick={() => onSelect(c)}>
                  <td className="px-3 py-3">
                    <div className="font-display ink text-sm" style={{ fontWeight: 700 }}>{c.venue}</div>
                    {c.notes && <div className="text-[10px] italic ocean truncate max-w-[200px]">{c.notes}</div>}
                  </td>
                  <td className="px-3 py-3">
                    <span className="text-[10px] font-display tracking-wider px-2 py-0.5 border" style={{ borderColor: c.channel === 'Private Sales' ? '#D78433' : '#006C90', color: c.channel === 'Private Sales' ? '#D78433' : '#006C90', fontWeight: 600 }}>
                      {c.channel === 'Private Sales' ? 'PRIVATE' : c.channel === 'Trade Retail' ? 'RETAIL' : '—'}
                    </span>
                  </td>
                  <td className="px-3 py-3 text-xs">
                    {(c.firstName || c.lastName) && <div className="ink">{c.firstName} {c.lastName}</div>}
                    {c.phone && <div className="ocean text-[10px]">{c.phone}</div>}
                    {c.email && <div className="ocean text-[10px] truncate max-w-[180px]">{c.email}</div>}
                  </td>
                  <td className="px-3 py-3 text-xs ink">{c.location}</td>
                  <td className="px-3 py-3 text-xs">
                    <span className="font-display tracking-wider ink" style={{ fontWeight: 700 }}>{c.accountManager === 'Unassigned' ? '—' : c.accountManager}</span>
                  </td>
                  <td className="px-3 py-3"><StatusBadge status={c.status} /></td>
                  <td className="px-3 py-3 text-xs ocean italic">{c.lastContacted || '—'}</td>
                  <td className="px-3 py-3 text-xs ink font-display" style={{ fontWeight: 700 }}>{c.totalSales > 0 ? ZAR(c.totalSales) : '—'}</td>
                  <td className="px-3 py-3"><ChevronRight className="w-4 h-4 copper" /></td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan="9" className="px-3 py-12 text-center italic ocean">No clients match your filters.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Mobile card list — hidden on desktop */}
      <div className="md:hidden space-y-2">
        {filtered.length === 0 && (
          <div className="premium-card p-6 text-center">
            <p className="italic ocean text-sm">No clients match your filters.</p>
          </div>
        )}
        {filtered.map(c => (
          <div key={c.id} className="premium-card p-4 active:bg-gold/10 cursor-pointer border-l-4 transition-colors"
            style={{ borderLeftColor: c.channel === 'Private Sales' ? '#D78433' : '#006C90' }}
            onClick={() => onSelect(c)}>
            <div className="flex items-start justify-between gap-2">
              <div className="flex-1 min-w-0">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="font-display text-sm ink truncate" style={{ fontWeight: 700 }}>{c.venue}</span>
                  <StatusBadge status={c.status} />
                </div>
                <p className="text-[11px] ocean italic mt-0.5 truncate">
                  {[c.firstName, c.lastName].filter(Boolean).join(' ')}
                  {c.location ? (c.firstName || c.lastName ? ' · ' : '') + c.location : ''}
                </p>
                <div className="flex items-center gap-3 mt-1.5 flex-wrap">
                  <span className="text-[9px] font-display tracking-wider px-1.5 py-0.5 border"
                    style={{ borderColor: c.channel === 'Private Sales' ? '#D78433' : '#006C90', color: c.channel === 'Private Sales' ? '#D78433' : '#006C90', fontWeight: 600 }}>
                    {c.channel === 'Private Sales' ? 'PRIVATE' : c.channel === 'Trade Retail' ? 'RETAIL' : '—'}
                  </span>
                  {c.accountManager && c.accountManager !== 'Unassigned' && (
                    <span className="text-[10px] font-display ink" style={{ fontWeight: 600 }}>{c.accountManager}</span>
                  )}
                  {c.lastContacted && <span className="text-[10px] ocean italic">{c.lastContacted}</span>}
                  {c.totalSales > 0 && <span className="text-[10px] font-display copper" style={{ fontWeight: 700 }}>{ZAR(c.totalSales)}</span>}
                </div>
              </div>
              <ChevronRight className="w-4 h-4 copper flex-shrink-0 mt-1" />
            </div>
            {c.notes && <p className="text-[10px] italic ocean mt-2 truncate">"{c.notes}"</p>}
          </div>
        ))}
      </div>
    </div>
  );
}

function FilterSelect({ label, value, onChange, options }) {
  return (
    <div className="relative">
      <span className="absolute left-3 top-1/2 -translate-y-1/2 font-display text-[10px] tracking-[0.2em] ocean pointer-events-none" style={{ fontWeight: 600 }}>{label.toUpperCase()}:</span>
      <select value={value} onChange={(e) => onChange(e.target.value)} className="w-full pl-16 pr-3 py-2.5 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper appearance-none cursor-pointer">
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
      <div className="pb-3 border-b border-ink/15">
        <p className="font-display text-[9px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>VISIT HISTORY</p>
        <div className="flex items-center justify-between gap-3 mt-1 flex-wrap">
          <h1 className="font-display text-2xl md:text-4xl ink" style={{ fontWeight: 700 }}>VISIT LOG</h1>
          <RepToggle active={filterRep} onChange={setFilterRep} />
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
                    {v.notes && <p className="text-[10px] italic ink/70 mt-1 truncate">"{v.notes}"</p>}
                  </div>
                  {/* Actions */}
                  <div className="flex items-center gap-1 flex-shrink-0">
                    <button type="button" onClick={() => onEdit(v)} className="p-2 hover:bg-ink/10 border border-ink/15">
                      <Edit2 className="w-3.5 h-3.5 ocean" />
                    </button>
                    <button type="button" onClick={() => onDelete(v.id)} className="p-2 hover:bg-red-50 border border-ink/15">
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
    <div className="fixed inset-0 z-50 flex items-end md:items-center justify-center md:p-4 bg-ink/70 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="bg-cream w-full md:max-w-3xl md:my-4 border-t-2 md:border-2 border-copper relative max-h-[95vh] md:max-h-none overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="absolute top-0 left-0 right-0 h-1 bg-copper z-10"></div>
        {/* Header */}
        <div className="p-4 md:p-6 relative overflow-hidden" style={{ background: '#003553' }}>
          <RaysBackdrop opacity={0.06} />
          <div className="flex items-center justify-between gap-4 relative">
            <div className="flex items-center gap-3">
              <AvanteLogo size={36} />
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
                <button key={r} type="button" onClick={() => { setSalesRep(r); setClientId(''); }} className={`py-3 font-display text-xs tracking-[0.2em] border transition-all ${salesRep === r ? 'bg-ink border-ink' : 'bg-cream border-ink/20 hover:border-copper'}`} style={{ color: salesRep === r ? '#FFFEF2' : '#003553', fontWeight: 700 }}>
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
              <input type="text" placeholder="Filter venues..." value={search} onChange={(e) => setSearch(e.target.value)} className="w-full pl-10 pr-3 py-2 border border-ink/20 bg-cream text-sm focus:outline-none focus:border-copper" />
            </div>
            <select value={clientId} onChange={(e) => setClientId(e.target.value)} className="w-full px-3 py-3 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper">
              <option value="">— Select a venue —</option>
              {repClients.map(c => (
                <option key={c.id} value={c.id}>{c.venue}{c.location ? ' · ' + c.location : ''}</option>
              ))}
            </select>
            {selectedClient && (
              <div className="mt-2 p-3 bg-gold/10 border-l-2 border-copper text-xs">
                <p className="ink font-display" style={{ fontWeight: 700 }}>{selectedClient.venue}</p>
                <p className="ocean italic">{selectedClient.location} · {selectedClient.channel || 'No channel'} · {selectedClient.status}</p>
              </div>
            )}
          </div>

          {/* Date + Outcome — side by side on all screens */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>DATE</label>
              <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="w-full px-3 py-3 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>OUTCOME</label>
              <select value={outcome} onChange={(e) => setOutcome(e.target.value)} className="w-full px-3 py-3 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper">
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
              <div className="relative">
                <button type="button" onClick={() => setSkuPickerOpen(!skuPickerOpen)} className="flex items-center gap-1.5 text-[10px] font-display tracking-[0.2em] copper" style={{ fontWeight: 700 }}>
                  <Plus className="w-3.5 h-3.5" /> ADD SKU
                </button>
                {skuPickerOpen && (
                  <div className="fixed md:absolute right-0 md:right-0 bottom-0 md:bottom-auto md:top-full md:mt-1 left-0 md:left-auto w-full md:w-80 bg-cream border-t-2 md:border-2 border-copper shadow-2xl z-30 max-h-[60vh] md:max-h-80 overflow-y-auto scrollbar-thin">
                    <div className="sticky top-0 bg-ink px-3 py-2.5 flex items-center justify-between">
                      <p className="font-display text-[10px] tracking-[0.25em]" style={{ color: '#FDB940', fontWeight: 600 }}>AVANTE CATALOGUE</p>
                      <button type="button" onClick={() => setSkuPickerOpen(false)} style={{ color: '#FFFEF2' }}><X className="w-4 h-4" /></button>
                    </div>
                    {SKU_CATALOGUE.map(sku => {
                      const inOrder = items.find(it => it.skuId === sku.id);
                      return (
                        <button key={sku.id} type="button" onClick={() => addSku(sku)}
                          className="w-full text-left px-4 py-3 hover:bg-gold/20 border-b border-ink/10 flex items-center justify-between active:bg-gold/30">
                          <div>
                            <p className="ink text-sm font-display" style={{ fontWeight: 700 }}>{sku.name}</p>
                            <p className="ocean text-xs italic">{ZAR(sku.price)} per unit</p>
                          </div>
                          {inOrder ? <span className="text-[9px] copper font-display tracking-wider" style={{ fontWeight: 700 }}>IN ORDER</span> : <Plus className="w-4 h-4 copper" />}
                        </button>
                      );
                    })}
                  </div>
                )}
              </div>
            </div>

            {items.length === 0 ? (
              <div className="border-2 border-dashed border-ink/20 p-5 text-center bg-cream">
                <p className="text-xs italic ocean">No SKUs added. Tap "ADD SKU" to start building the order.</p>
              </div>
            ) : (
              <div className="border border-ink/20">
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
                              className="w-full px-2 py-2 border border-ink/20 bg-cream text-sm text-center focus:outline-none focus:border-copper" />
                          </div>
                          <div>
                            <p className="font-display text-[9px] tracking-[0.15em] ocean mb-1" style={{ fontWeight: 600 }}>UNIT R</p>
                            <input type="number" min="0" step="0.01" value={it.unitPrice}
                              onChange={(e) => updateItem(it.skuId, 'unitPrice', e.target.value)}
                              className="w-full px-2 py-2 border border-ink/20 bg-cream text-sm text-right focus:outline-none focus:border-copper"
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
                  <div className="grid grid-cols-12 gap-2 px-3 py-2 bg-ink/5 border-b border-ink/10">
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
                      <div key={it.skuId} className="grid grid-cols-12 gap-2 px-3 py-2 border-b border-ink/10 items-center">
                        <div className="col-span-5">
                          <p className="ink text-xs font-display" style={{ fontWeight: 700 }}>{it.name}</p>
                          {discounted && <p className="text-[9px] copper italic">from {ZAR(it.listPrice)}</p>}
                        </div>
                        <div className="col-span-2">
                          <input type="number" min="0" step="1" value={it.qty}
                            onChange={(e) => updateItem(it.skuId, 'qty', e.target.value)}
                            className="w-full px-2 py-1 border border-ink/20 bg-cream text-xs text-center focus:outline-none focus:border-copper" />
                        </div>
                        <div className="col-span-2">
                          <input type="number" min="0" step="0.01" value={it.unitPrice}
                            onChange={(e) => updateItem(it.skuId, 'unitPrice', e.target.value)}
                            className="w-full px-2 py-1 border border-ink/20 bg-cream text-xs text-right focus:outline-none focus:border-copper"
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
            <textarea value={notes} onChange={(e) => setNotes(e.target.value)} rows="3" placeholder="What happened in the meeting? Key decision-makers, objections, opportunities..." className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper resize-none" />
          </div>

          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>FOLLOW-UP NOTES</label>
            <textarea value={followUp} onChange={(e) => setFollowUp(e.target.value)} rows="2" placeholder="Next action, due date, who owns it..." className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper resize-none" />
          </div>

          {/* Actions */}
          {validationError && (
            <div className="px-3 py-2 border-l-4 text-xs" style={{ borderLeftColor: '#9c2c2c', background: 'rgba(156,44,44,0.08)', color: '#9c2c2c' }}>
              <span className="font-display tracking-wider" style={{ fontWeight: 700 }}>ERROR:</span> {validationError}
            </div>
          )}
          {/* Mobile: stacked full-width buttons. Desktop: row layout */}
          <div className="pt-4 border-t border-ink/10 space-y-2 md:space-y-0 md:flex md:items-center md:justify-between md:gap-3 md:flex-wrap">
            <button type="button" onClick={handleEmailOrder}
              className="w-full md:w-auto flex items-center justify-center gap-2 px-4 py-3 font-display text-xs tracking-[0.25em] border"
              style={{ color: '#006C90', fontWeight: 700, borderColor: '#006C90' }}>
              <Mail className="w-4 h-4" /> EMAIL ORDER
            </button>
            <div className="flex items-center gap-2">
              <button type="button" onClick={onClose} className="flex-1 md:flex-none px-5 py-3 font-display text-xs tracking-[0.25em] ink border border-ink/20" style={{ fontWeight: 700 }}>CANCEL</button>
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
    <div className="fixed inset-0 z-[70] flex items-start md:items-center justify-center p-4 bg-ink/80 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
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
            <div className="flex items-center gap-3 p-3 bg-ink/5 border border-ink/10">
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
              className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper resize-none"
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
            <p className="text-xs ink p-3 bg-ink/5 border border-ink/10 italic">{composed.subject}</p>
          </div>

          {/* Order summary */}
          <div className="grid grid-cols-3 gap-3 py-3 border-y border-ink/10">
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
              <pre className="mt-2 p-3 bg-ink/5 border border-ink/10 text-[11px] ink whitespace-pre-wrap font-mono max-h-72 overflow-y-auto scrollbar-thin">{composed.body}</pre>
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
                className="flex items-center justify-center gap-2 px-4 py-3 font-display text-xs tracking-[0.2em] hover:opacity-90 no-underline border border-ink/20"
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
                className="flex items-center justify-center gap-2 px-4 py-3 font-display text-xs tracking-[0.2em] hover:opacity-90 no-underline border border-ink/20"
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
          <button type="button" onClick={onClose} className="px-5 py-2.5 font-display text-xs tracking-[0.25em] ink hover:bg-ink/5" style={{ fontWeight: 700 }}>
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
      className="flex items-center gap-2 px-4 py-2.5 font-display text-[10px] tracking-[0.25em] border border-ink/20 hover:bg-ink/5"
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

  const handleCreate = () => {
    console.log('[NewClient] create clicked', form);
    if (!form.venue.trim()) { alert('Venue name is required.'); return; }
    try {
      onSave(form);
    } catch (err) {
      console.error('[NewClient] save threw', err);
      alert('Could not create client: ' + (err?.message || err));
    }
  };

  const setF = (k, v) => setForm({ ...form, [k]: v });

  return (
    <div className="fixed inset-0 z-[60] flex items-start md:items-center justify-center p-4 bg-ink/80 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="bg-cream max-w-2xl w-full my-4 border-2 border-copper" onClick={(e) => e.stopPropagation()}>
        <div className="p-5 flex items-center justify-between" style={{ background: '#003553' }}>
          <div className="flex items-center gap-3">
            <UserPlus className="w-5 h-5" style={{ color: '#FDB940' }} />
            <div>
              <p className="font-display text-[10px] tracking-[0.4em]" style={{ color: '#FDB940', fontWeight: 600 }}>NEW CLIENT</p>
              <h2 className="font-display text-xl" style={{ color: '#FFFEF2', fontWeight: 700, letterSpacing: '0.08em' }}>ADD TO {defaultRep.toUpperCase()}'S BOOK</h2>
            </div>
          </div>
          <button type="button" onClick={onClose} className="p-2 hover:opacity-70" style={{ color: '#FFFEF2' }}><X className="w-5 h-5" /></button>
        </div>
        <div className="p-6 space-y-4">
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>VENUE NAME *</label>
            <input type="text" value={form.venue} onChange={(e) => setF('venue', e.target.value)} placeholder="e.g. The Grand Hotel" className="w-full px-3 py-2.5 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper" autoFocus />
          </div>
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>FIRST NAME</label>
              <input type="text" value={form.firstName} onChange={(e) => setF('firstName', e.target.value)} className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>LAST NAME</label>
              <input type="text" value={form.lastName} onChange={(e) => setF('lastName', e.target.value)} className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>EMAIL</label>
              <input type="email" value={form.email} onChange={(e) => setF('email', e.target.value)} className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>PHONE</label>
              <input type="tel" value={form.phone} onChange={(e) => setF('phone', e.target.value)} className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>LOCATION</label>
              <input type="text" value={form.location} onChange={(e) => setF('location', e.target.value)} placeholder="e.g. Camps Bay" className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>DISTRIBUTOR</label>
              <input type="text" value={form.distributor} onChange={(e) => setF('distributor', e.target.value)} className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper" />
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>CHANNEL</label>
              <select value={form.channel} onChange={(e) => setF('channel', e.target.value)} className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper">
                {CHANNELS.map(c => <option key={c}>{c}</option>)}
              </select>
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>LEAD SOURCE</label>
              <select value={form.leadSource} onChange={(e) => setF('leadSource', e.target.value)} className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper">
                {LEAD_SOURCES.map(s => <option key={s}>{s}</option>)}
              </select>
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>PRIORITY</label>
              <select value={form.priority} onChange={(e) => setF('priority', e.target.value)} className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper">
                {PRIORITIES.map(p => <option key={p}>{p}</option>)}
              </select>
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>ACCOUNT MANAGER</label>
              <select value={form.accountManager} onChange={(e) => setF('accountManager', e.target.value)} className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper">
                {SALES_REPS.map(r => <option key={r}>{r}</option>)}
              </select>
            </div>
          </div>
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>NOTES</label>
            <textarea value={form.notes} onChange={(e) => setF('notes', e.target.value)} rows="2" className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper resize-none" />
          </div>
          <div className="flex justify-end gap-3 pt-3 border-t border-ink/10">
            <button type="button" onClick={onClose} className="px-5 py-2.5 font-display text-xs tracking-[0.25em] ink hover:bg-ink/5" style={{ fontWeight: 700 }}>CANCEL</button>
            <button type="button" onClick={handleCreate} className="bg-copper hover:bg-gold px-6 py-2.5 font-display text-xs tracking-[0.25em] flex items-center gap-2" style={{ color: '#FFFEF2', fontWeight: 700 }}>
              <UserPlus className="w-4 h-4" /> CREATE CLIENT
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
    <div className="fixed inset-0 z-50 flex items-end md:items-start justify-center md:p-4 bg-ink/70 backdrop-blur-sm overflow-y-auto" onClick={onClose}>
      <div className="bg-cream w-full md:max-w-3xl md:my-4 border-t-2 md:border-2 border-copper max-h-[95vh] md:max-h-none overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="p-4 md:p-6 relative overflow-hidden" style={{ background: '#003553' }}>
          <RaysBackdrop opacity={0.06} />
          <div className="flex items-start justify-between gap-4 relative">
            <div className="flex items-start gap-3">
              <AvanteLogo size={36} />
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
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 pb-4 border-b border-ink/10">
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
              <textarea value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} rows="3" className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper resize-none" />
            ) : (
              <p className="text-sm ink italic">{client.notes || '—'}</p>
            )}
          </div>

          {/* === PURCHASE HISTORY (SKUs aggregated) === */}
          {skuTotals.length > 0 && (
            <div className="pt-5 border-t border-ink/10">
              <div className="flex items-center justify-between mb-3">
                <p className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>PURCHASE HISTORY · SKU BREAKDOWN</p>
                <span className="text-[10px] italic ocean">All-time totals across {visits.filter(v => v.items && v.items.length > 0).length} order{visits.filter(v => v.items && v.items.length > 0).length === 1 ? '' : 's'}</span>
              </div>
              <div className="border border-ink/15">
                <div className="grid grid-cols-12 gap-2 px-3 py-2 bg-ink/5 border-b border-ink/10">
                  <div className="col-span-6 font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>SKU</div>
                  <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-center" style={{ fontWeight: 600 }}>UNITS</div>
                  <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-right" style={{ fontWeight: 600 }}>REVENUE</div>
                  <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-right" style={{ fontWeight: 600 }}>LAST</div>
                </div>
                {skuTotals.map(s => (
                  <div key={s.skuId} className="grid grid-cols-12 gap-2 px-3 py-2 border-b border-ink/10 items-center text-xs">
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
            <div className="pt-5 border-t border-ink/10">
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
                      <div className={`border ${isLatest ? 'border-copper' : isHistorical ? 'border-ink/10 border-dashed' : 'border-ink/15'} bg-cream`}>
                        <div className="flex items-baseline justify-between gap-2 px-3 py-2 border-b border-ink/10" style={{ background: isLatest ? 'rgba(253,185,64,0.08)' : isHistorical ? 'rgba(0,108,144,0.04)' : 'transparent' }}>
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
                                    <div key={i} className="flex items-center justify-between text-[11px] px-2 py-1 bg-ink/5">
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
            <div className="pt-5 border-t border-ink/10">
              <p className="font-display text-[10px] tracking-[0.3em] copper mb-3" style={{ fontWeight: 600 }}>CALL CYCLE · VISIT TIMELINE</p>
              <div className="border border-dashed border-ink/20 p-6 text-center">
                <p className="text-xs italic ocean">No visits logged for this client yet.</p>
                <p className="text-xs italic ocean mt-1">Use "Log Visit" to start tracking the call cycle.</p>
              </div>
            </div>
          )}

          <div className="flex items-center justify-end gap-3 pt-4 border-t border-ink/10">
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
        <input type="text" value={value || ''} onChange={(e) => onChange(e.target.value)} className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper" />
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
        <select value={value || ''} onChange={(e) => onChange(e.target.value)} className="w-full px-3 py-2 border border-ink/20 bg-cream font-body text-sm focus:outline-none focus:border-copper">
          {options.map(o => <option key={o} value={o}>{o}</option>)}
        </select>
      ) : (
        <div className="text-sm ink py-1">{value || <span className="italic ocean">—</span>}</div>
      )}
    </div>
  );
}
