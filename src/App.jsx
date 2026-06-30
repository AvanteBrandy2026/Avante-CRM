import React, { useState, useEffect, useMemo, useRef } from 'react';
import { LayoutDashboard, Users, ClipboardList, Settings, TrendingUp, Phone, Mail, Search, Plus, X, ChevronRight, DollarSign, Award, Activity, Briefcase, Wine, ArrowUpRight, Save, RotateCcw, Target, BarChart3, Trash2, Download, FileSpreadsheet, UserPlus, Edit2, ShoppingCart, Package, ChevronDown, Bell, LogOut } from 'lucide-react';
import * as XLSX from 'xlsx';

// ── USER AUTH CONFIG ─────────────────────────────────────────────────────────
// Simple credential map — internal tool, no external auth service needed.
// Each user maps to a rep name, role, and display name.
const USERS = {
  'alex@breakfreebeverages.com':     { password: 'alex123',     rep: 'Alex',    role: 'rep',     name: 'Alex' },
  'lehmarc@breakfreebeverages.com':  { password: 'lehmarc123',  rep: 'Lehmarc', role: 'rep',     name: 'Lehmarc' },
  'lonwabo@breakfreebeverages.com':  { password: 'lonwabo123',  rep: 'Loydz',   role: 'rep',     name: 'Loydz' },
  'louis@breakfreebeverages.com':    { password: 'louis123',    rep: 'Louis',   role: 'rep',     name: 'Louis' },
  'anthony@breakfreebeverages.com':  { password: 'anthony123',  rep: 'Anthony', role: 'rep',     name: 'Anthony' },
  'matthew@breakfreebeverages.com':  { password: 'matthew123',  rep: 'Matthew', role: 'manager', name: 'Matthew' },
};
const AUTH_KEY = 'avante_crm_user';
const getStoredUser = () => { try { return JSON.parse(localStorage.getItem(AUTH_KEY)); } catch { return null; } };
const storeUser = (u) => { try { localStorage.setItem(AUTH_KEY, JSON.stringify(u)); } catch {} };
const clearUser = () => { try { localStorage.removeItem(AUTH_KEY); } catch {} };
const isManager = (u) => u?.role === 'manager';

// ── LOGIN SCREEN ─────────────────────────────────────────────────────────────

const SEED_CLIENTS = [{"id":1,"venue":"The Lawns","channel":"B2B","firstName":"","lastName":"","location":"Camps Bay","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Low","lastContacted":"","status":"New","notes":"Need to make contact","totalSales":0},{"id":2,"venue":"50 on Gugs","channel":"B2B","firstName":"Sinbad","lastName":"","location":"Gugulethu","distributor":"Ultra","email":"","phone":"082 553 3599","leadSource":"Referral","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-18","status":"New","notes":"","totalSales":0},{"id":3,"venue":"Mitchell's","channel":"B2B","firstName":"Leroy","lastName":"","location":"CBD","distributor":"","email":"","phone":"660793944","leadSource":"Walk into Store","accountManager":"Alex","priority":"Medium","lastContacted":"2026-01-29","status":"New","notes":"Parking for now","totalSales":0},{"id":4,"venue":"Anthm","channel":"B2B","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":5,"venue":"Athletic Club & Social","channel":"B2B","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":6,"venue":"Hennies","channel":"B2B","firstName":"","lastName":"","location":"Stellenbosch","distributor":"","email":"","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"Medium","lastContacted":"2026-02-10","status":"New","notes":"Parking for now","totalSales":0},{"id":7,"venue":"Beyond","channel":"B2B","firstName":"Sebastian","lastName":"Stehr","location":"Constantia","distributor":"","email":"chef@beyondrestaurant.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-10","status":"New","notes":"Follow up in June","totalSales":0},{"id":8,"venue":"Rust en Vrede","channel":"B2B","firstName":"Fabio","lastName":"Daniel","location":"Stellenbosch","distributor":"","email":"fabio@rustenvrede.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-10","status":"Converted","notes":"Rejected","totalSales":0},{"id":9,"venue":"Blue Peter","channel":"B2B","firstName":"Marshal","lastName":"","location":"Blouberg","distributor":"","email":"","phone":"083 981 8770","leadSource":"Referral","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"Converted","notes":"","totalSales":0},{"id":10,"venue":"The Tasting Room at Creation","channel":"B2B","firstName":"Eleanor","lastName":"Niehaus","location":"Hermanus","distributor":"","email":"eleanor@creationwines.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-10","status":"New","notes":"Rejected","totalSales":0},{"id":11,"venue":"Dusk","channel":"B2B","firstName":"Callen","lastName":"Austin","location":"Stellenbosch","distributor":"","email":"","phone":"081 409 6544","leadSource":"Referral","accountManager":"Alex","priority":"Low","lastContacted":"2026-03-11","status":"Contacted","notes":"Rejected","totalSales":0},{"id":12,"venue":"Firemans Arms","channel":"B2B","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-20","status":"New","notes":"Promo finished. Not a focus.","totalSales":0},{"id":13,"venue":"Chefs Table","channel":"B2B","firstName":"Karl","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"2026-04-20","status":"Converted","notes":"Calling today","totalSales":0},{"id":14,"venue":"Bossa Group","channel":"B2B","firstName":"Chris","lastName":"","location":"Paarl","distributor":"","email":"","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-03-24","status":"Prospect","notes":"Parking for now","totalSales":0},{"id":15,"venue":"Chefs Warehouse Bree","channel":"B2B","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":16,"venue":"Aubergine","channel":"B2B","firstName":"Harald","lastName":"Bresselschmidt","location":"CBD","distributor":"","email":"chef@aubergine.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"Converted","notes":"Meeting on 24th","totalSales":0},{"id":17,"venue":"Chorus","channel":"B2B","firstName":"Mareli","lastName":"Basson","location":"Somerset West","distributor":"","email":"mareli@bertusbasson.com / info@bertusbasson.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Calling today","totalSales":0},{"id":18,"venue":"Club Heritage","channel":"B2B","firstName":"Pholoso","lastName":"","location":"","distributor":"","email":"","phone":"073 798 1918","leadSource":"Referral","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-17","status":"Converted","notes":"","totalSales":0},{"id":19,"venue":"Club Soho","channel":"B2B","firstName":"Krys","lastName":"","location":"","distributor":"","email":"djkrys86@gmail.com","phone":"083 533 9789","leadSource":"Referral","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"New","notes":"","totalSales":0},{"id":20,"venue":"Foxcroft","channel":"B2B","firstName":"Zoe","lastName":"Wepener","location":"","distributor":"","email":"beauadmin@chefswarehouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Emailed. Waiting for response.","totalSales":0},{"id":21,"venue":"Cyra","channel":"B2B","firstName":"Candice","lastName":"Philip","location":"Jo'Burg","distributor":"","email":"chefcandicephilip@gmail.com","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-03-10","status":"New","notes":"","totalSales":0},{"id":22,"venue":"Gigi","channel":"B2B","firstName":"Moses","lastName":"Moloi","location":"CBD","distributor":"","email":"moses@gigirestaurant.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"No response. Call.","totalSales":0},{"id":23,"venue":"Grand Cafe","channel":"B2B","firstName":"Rees","lastName":"Maxwell","location":"Sea Point","distributor":"","email":"","phone":"082 874 2299","leadSource":"Referral","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"Prospect","notes":"Struggles to get hold of Rees. Will keep persisting","totalSales":0},{"id":24,"venue":"Grub & Vne","channel":"B2B","firstName":"Matthew","lastName":"Manning","location":"","distributor":"","email":"info@mattmanningchef.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Emailed. Waiting for response.","totalSales":0},{"id":25,"venue":"H\u014dseki","channel":"B2B","firstName":"Virgil","lastName":"Kahn","location":"Stellenbosch","distributor":"","email":"lodge.chef@delaire.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"No response. Call.","totalSales":0},{"id":26,"venue":"La Colombe","channel":"B2B","firstName":"James","lastName":"Gaag","location":"Hout Bay","distributor":"","email":"james@lacolombe.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":27,"venue":"Post & Pepper","channel":"B2B","firstName":"Jess","lastName":"van Dyk","location":"Stellenbosch","distributor":"","email":"jess@postandpepper.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"Converted","notes":"Follow up in June","totalSales":0},{"id":28,"venue":"Serendipity","channel":"B2B","firstName":"Lizelle","lastName":"Stolze","location":"Wiilderness","distributor":"","email":"chef123@mweb.co.za/\nrestaurant@serendipitywilderness.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Call.","totalSales":0},{"id":29,"venue":"aha Hotels & Lodges","channel":"B2B","firstName":"Ruzandri","lastName":"","location":"","distributor":"","email":"ruzandri.stoltz@aha.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":30,"venue":"Hyatt Regency Cape Town","channel":"B2B","firstName":"Sasha","lastName":"Lewis","location":"","distributor":"","email":"sasha.lewis@hyatt.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":31,"venue":"Fac Cactus - Blouberg","channel":"B2B","firstName":"Lesley","lastName":"","location":"Blouberg","distributor":"","email":"simon@fatcactus.co.za","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"High","lastContacted":"2026-01-22","status":"New","notes":"","totalSales":0},{"id":32,"venue":"Silo (RP)","channel":"B2B","firstName":"Gerrit","lastName":"","location":"","distributor":"","email":"gerrit@trp.travel","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":33,"venue":"Steenberg GC","channel":"B2B","firstName":"Tatiana","lastName":"","location":"","distributor":"","email":"fbmanager@steenberggolfclub.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":34,"venue":"Twelve Apostles Hotel","channel":"B2B","firstName":"Mo","lastName":"Taliet","location":"","distributor":"","email":"groupproman12a@rchmail.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":35,"venue":"Westlake GC","channel":"B2B","firstName":"Tina","lastName":"","location":"","distributor":"","email":"events@westlakegolfclub.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":36,"venue":"De Zalze GC","channel":"B2B","firstName":"Ian","lastName":"","location":"","distributor":"","email":"FBManager@dezalzegolf.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":37,"venue":"Pearl Valley GC","channel":"B2B","firstName":"Stafford","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":38,"venue":"Arabella GC","channel":"B2B","firstName":"Sterlia","lastName":"Van Der Merwe","location":"","distributor":"","email":"sterlia.vandermerwe@southernsun.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":39,"venue":"Red Carnation Hotels","channel":"B2B","firstName":"Jameel","lastName":"","location":"","distributor":"","email":"foodbeverage@oysterbox.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":40,"venue":"Royal Johannesburg GC","channel":"B2B","firstName":"Erik","lastName":"","location":"","distributor":"","email":"fb@royaljhb.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":41,"venue":"Happy Folks","channel":"B2B","firstName":"Jacobus","lastName":"","location":"Stellenbosch / Blouberg","distributor":"","email":"blouberg@happyfolks.co.za","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"Low","lastContacted":"2026-01-20","status":"New","notes":"Need to send him a email and get Franchise approval to list Avante.","totalSales":0},{"id":42,"venue":"Madame Zingara Group","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":43,"venue":"More Family Collection","channel":"B2B","firstName":"Lara","lastName":"","location":"","distributor":"","email":"lara@more.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"Follow up","totalSales":0},{"id":44,"venue":"The Bay Hotel","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":45,"venue":"Premier Hotels","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":46,"venue":"Kaapstadt Brauhaus","channel":"B2B","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-21","status":"New","notes":"","totalSales":0},{"id":47,"venue":"Paarl Brandy Bar","channel":"B2B","firstName":"Allan","lastName":"","location":"Paarl","distributor":"","email":"","phone":"082 447 2538","leadSource":"Event","accountManager":"Alex","priority":"Low","lastContacted":"2026-04-02","status":"Qualified","notes":"Opening bar in Feb. Need to set up eeting this week","totalSales":0},{"id":48,"venue":"De Tafel","channel":"B2B","firstName":"Gregory","lastName":"Henderson","location":"Wynberg","distributor":"","email":"execchef@palmhouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"New","notes":"Waiting on response","totalSales":0},{"id":49,"venue":"Delaire Graff Restaurant","channel":"B2B","firstName":"Clinton","lastName":"Jacobs","location":"Stellenbosch","distributor":"","email":"dgr.headchef@delaire.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"New","notes":"No response. Call.","totalSales":0},{"id":50,"venue":"Fable","channel":"B2B","firstName":"Josh","lastName":"Sarembock","location":"CBD","distributor":"","email":"","phone":"060 313 4539","leadSource":"Event","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"Prospect","notes":"Keen to stock. Contracts end in June.","totalSales":0},{"id":51,"venue":"Fermier","channel":"B2B","firstName":"Adriaan","lastName":"Maree","location":"Pretoria","distributor":"","email":"bookings@fermierrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-02","status":"New","notes":"No response. Call.","totalSales":0},{"id":52,"venue":"Salsify at the Roundhouse","channel":"B2B","firstName":"Ryan","lastName":"Cole","location":"Camps Bay","distributor":"","email":"ryan@salsify.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Figuring out distribution","totalSales":0},{"id":53,"venue":"Les Cr\u00e9atifs","channel":"B2B","firstName":"Wandile","lastName":"Mabaso","location":"Jo'Burg","distributor":"","email":"chefwmabaso@gmail.com","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-10","status":"Converted","notes":"Meeting end of May","totalSales":0},{"id":54,"venue":"Terrarium","channel":"B2B","firstName":"Anlou","lastName":"Erasmus","location":"Waterfront","distributor":"","email":"qvchef@queenvictoriahotel.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":55,"venue":"The Chefs' Table","channel":"B2B","firstName":"Mathew","lastName":"Armbruster","location":"","distributor":"","email":"mathew@thechefstable.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":56,"venue":"The Jordan Restaurant with Marthinus Ferreira","channel":"B2B","firstName":"Leigh","lastName":"Simmidari","location":"Stellenbosch","distributor":"","email":"restaurant@jordanwines.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"New","notes":"Calling today","totalSales":0},{"id":57,"venue":"Upper Union","channel":"B2B","firstName":"Amori","lastName":"Burger","location":"CBD","distributor":"","email":"amori@upperunion.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-03","status":"Converted","notes":"Followed up. Waiting.","totalSales":0},{"id":58,"venue":"Chefs Warehouse Maison","channel":"B2B","firstName":"David Schneider","lastName":"Xavier Francis (head chef)","location":"Franschoek","distributor":"","email":"david@chefswarehouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-06","status":"New","notes":"No response. Call.","totalSales":0},{"id":59,"venue":"Medusa","channel":"B2B","firstName":"Vuyani","lastName":"","location":"Milnerton","distributor":"","email":"medusaloungerooftop@gmail.com","phone":"072 258 2463","leadSource":"Cold Call","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-16","status":"Converted","notes":"Done Deals","totalSales":0},{"id":60,"venue":"MIlnerton Golf Club","channel":"B2B","firstName":"Alfie","lastName":"Schneeburger","location":"Milnerton","distributor":"","email":"catering@milgolf.co.za","phone":"082 576 7887","leadSource":"Event","accountManager":"Loydz","priority":"Low","lastContacted":"2026-01-20","status":"New","notes":"","totalSales":0},{"id":61,"venue":"Misthom Lounge","channel":"B2B","firstName":"Lee","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"Low","lastContacted":"2026-02-20","status":"New","notes":"","totalSales":0},{"id":62,"venue":"Kove Collection","channel":"B2B","firstName":"Benito","lastName":"","location":"Camps Bay","distributor":"Karino","email":"","phone":"063 212 7038","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-07","status":"Prospect","notes":"Benito likes Avante. We have a relationship. Waiting for an order and which restaurants.","totalSales":0},{"id":63,"venue":"Mozambik","channel":"B2B","firstName":"Uhuru","lastName":"","location":"","distributor":"","email":"","phone":"084 057 0572","leadSource":"Walk into Store","accountManager":"Loydz","priority":"High","lastContacted":"2026-01-21","status":"New","notes":"","totalSales":0},{"id":64,"venue":"Noah","channel":"B2B","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":65,"venue":"Eike","channel":"B2B","firstName":"Michael","lastName":"Fuller","location":"Stellenbosch","distributor":"","email":"kitchen@eikerestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"Callling today","totalSales":0},{"id":66,"venue":"FABER","channel":"B2B","firstName":"Dale","lastName":"Stevens","location":"Paarl","distributor":"","email":"faber@avondalewine.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"No response. Call.","totalSales":0},{"id":67,"venue":"TTK Fledgelings","channel":"B2B","firstName":"Nathan","lastName":"Clarke","location":"CBD","distributor":"","email":"steven@lukedaleroberts.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":68,"venue":"Life and Brand Portfolo","channel":"B2B","firstName":"Jason","lastName":"","location":"Cape Town","distributor":"","email":"jason@lifeandbrand.co.za","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-08","status":"New","notes":"Waiting for response","totalSales":0},{"id":69,"venue":"Spice Route","channel":"B2B","firstName":"Gigi","lastName":"Bisogno","location":"Paarl","distributor":"","email":"","phone":"083 441 9870","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-08","status":"Contacted","notes":"","totalSales":0},{"id":70,"venue":"Qunu","channel":"B2B","firstName":"Matthew","lastName":"Foxon","location":"Jo'Burg","distributor":"","email":"matthewf@saxon.co.za/scottd@saxon.co.za","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-10","status":"New","notes":"Given to Loydz","totalSales":0},{"id":71,"venue":"Room13, Savage & Concept","channel":"B2B","firstName":"Prince","lastName":"","location":"","distributor":"","email":"","phone":"061 982 6349","leadSource":"Referral","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"New","notes":"","totalSales":0},{"id":72,"venue":"La Petite Colombe","channel":"B2B","firstName":"Peter","lastName":"Duncan","location":"Franschoek","distributor":"","email":"peter@lapetitecolombe.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-08","status":"Lost","notes":"Revisit in Q2","totalSales":0},{"id":73,"venue":"Le coin Fran\u00e7ais","channel":"B2B","firstName":"Darren","lastName":"Badenhorst","location":"Franschoek","distributor":"","email":"darren@lecoinfrancais.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-08","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":74,"venue":"PIER","channel":"B2B","firstName":"John","lastName":"Norris-Rogers","location":"Waterfront","distributor":"","email":"john@pier.restaurant","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-08","status":"New","notes":"Callling today","totalSales":0},{"id":75,"venue":"Seat - Blouberg","channel":"B2B","firstName":"Anton","lastName":"","location":"Blouberg","distributor":"","email":"info@seatrestaurant.com","phone":"","leadSource":"Walk into Store","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-01-22","status":"New","notes":"","totalSales":0},{"id":76,"venue":"LifeGrande Group","channel":"B2B","firstName":"Jeanel / John","lastName":"","location":"","distributor":"","email":"jeanel@lifegrandcafe.com / john@lifegrandcafe.com","phone":"","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-09","status":"Prospect","notes":"Waiting for response","totalSales":0},{"id":77,"venue":"Stellenbosch GC","channel":"B2B","firstName":"Gerhard","lastName":"","location":"Stellenbosch","distributor":"Karino","email":"chef@stbgolf.com","phone":"072 198 7912","leadSource":"Event","accountManager":"Alex","priority":"High","lastContacted":"2026-04-09","status":"Prospect","notes":"Waitng for order coonfirmation","totalSales":0},{"id":78,"venue":"Galjoen","channel":"B2B","firstName":"Isca","lastName":"Stoltz","location":"CBD","distributor":"","email":"eat@galjoencpt.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-09","status":"New","notes":"No response. Call.","totalSales":0},{"id":79,"venue":"UCT Bar","channel":"B2B","firstName":"Rob","lastName":"Munroe","location":"UCT","distributor":"","email":"","phone":"079 569 4688","leadSource":"Referral","accountManager":"Alex","priority":"High","lastContacted":"2026-04-13","status":"Converted","notes":"Confirmed","totalSales":0},{"id":80,"venue":"SoHo","channel":"B2B","firstName":"DJ Krys","lastName":"","location":"Long Street","distributor":"","email":"083 533 9789","phone":"083 533 9789","leadSource":"Referral","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-21","status":"Converted","notes":"","totalSales":0},{"id":81,"venue":"Belly of the Beast","channel":"B2B","firstName":"Anouchka","lastName":"Horn","location":"CBD","distributor":"","email":"eat@bellyofthebeast.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-13","status":"New","notes":"No response. Call.","totalSales":0},{"id":82,"venue":"Cavalli","channel":"B2B","firstName":"Lucas","lastName":"Carstens","location":"Stellenbosch","distributor":"","email":"chef@cavalliestate.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-13","status":"New","notes":"Calling today","totalSales":0},{"id":83,"venue":"Embarc","channel":"B2B","firstName":"Darren","lastName":"O'Donovan","location":"Jo'Burg","distributor":"","email":"info@embarcrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-13","status":"New","notes":"No response. Call.","totalSales":0},{"id":84,"venue":"Thirsty Scarecrow","channel":"B2B","firstName":"Ivan","lastName":"Botha","location":"Stellenbosch","distributor":"Direct","email":"kzetler@gmail.com","phone":"084 263 4243","leadSource":"Event","accountManager":"Alex","priority":"Low","lastContacted":"2026-04-14","status":"Converted","notes":"Need to drop barrels. Waiting for order amount.","totalSales":0},{"id":85,"venue":"&Beyond","channel":"B2B","firstName":"Nicole Robinson","lastName":"","location":"","distributor":"","email":"nicole.robinson@andbeyond.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-14","status":"New","notes":"","totalSales":0},{"id":86,"venue":"Talking to Strangers","channel":"B2B","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":87,"venue":"Wolfgat","channel":"B2B","firstName":"Kobus","lastName":"van der Merwe","location":"","distributor":"","email":"jjvandermerwe@gmail.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-15","status":"New","notes":"Calling today","totalSales":0},{"id":88,"venue":"Thornybush","channel":"B2B","firstName":"Zanele","lastName":"","location":"","distributor":"","email":"","phone":"157931976","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-15","status":"New","notes":"","totalSales":0},{"id":89,"venue":"The Waterside","channel":"B2B","firstName":"Roxy","lastName":"Mudie","location":"Waterfront","distributor":"","email":"roxy@thewaterside.restaurant","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-16","status":"New","notes":"Calling today","totalSales":0},{"id":90,"venue":"Clovelly Golf Club","channel":"B2B","firstName":"Paul","lastName":"","location":"","distributor":"","email":"gm@clovelly.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-16","status":"New","notes":"","totalSales":0},{"id":91,"venue":"The LivingRoom","channel":"B2B","firstName":"Johannes","lastName":"Richter","location":"Durban","distributor":"","email":"info@summerhillkzn.com","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-01-22","status":"New","notes":"Given to Loydz","totalSales":0},{"id":92,"venue":"The Pot Luck Club, Johannesburg","channel":"B2B","firstName":"Ebie du Toit/ Vicky Peech","lastName":"Ebie du Toit/ Vicky Peech","location":"Jo'Burg","distributor":"","email":"ebi@plcj.co.za traveller@thepeech.co.za","phone":"","leadSource":"Email","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-01-22","status":"New","notes":"","totalSales":0},{"id":93,"venue":"Marriott / Protea Hotels","channel":"B2B","firstName":"Unathi","lastName":"Dyonase","location":"","distributor":"","email":"Unathi.dyonase@marriot.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-17","status":"New","notes":"","totalSales":0},{"id":94,"venue":"Chefs Warehouse Beau Constantia","channel":"B2B","firstName":"Ivor","lastName":"Jones","location":"Constantia","distributor":"","email":"ivor@chefswarehouse.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-20","status":"New","notes":"Followed up in March","totalSales":0},{"id":95,"venue":"COY","channel":"B2B","firstName":"Teenola","lastName":"Govender","location":"CBD","distributor":"","email":"teenola@coyrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-20","status":"New","notes":"Callling today","totalSales":0},{"id":96,"venue":"The Red Room","channel":"B2B","firstName":"Caroline","lastName":"Lamb","location":"CBD","distributor":"","email":"caroline@cwredroom.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-21","status":"Converted","notes":"Meeting next week","totalSales":0},{"id":97,"venue":"Tsogo Sun Mossel Bay","channel":"B2B","firstName":"George","lastName":"Webber","location":"Mossel Bay","distributor":"","email":"","phone":"838227550","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-21","status":"New","notes":"","totalSales":0},{"id":98,"venue":"\u00eblgr.","channel":"B2B","firstName":"Jesper","lastName":"Nillson","location":"CBD","distributor":"","email":"jesper@elgr.co.za","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-22","status":"New","notes":"Callling today","totalSales":0},{"id":99,"venue":"FYN","channel":"B2B","firstName":"Peter","lastName":"Templehoff","location":"CBD","distributor":"","email":"tempelhoff@gmail.com/ Kerry@fynrestaurant.com","phone":"","leadSource":"Email","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-22","status":"New","notes":"Followed up. Waiting.","totalSales":0},{"id":100,"venue":"Quay 4","channel":"B2B","firstName":"Leroy","lastName":"","location":"Waterfront","distributor":"","email":"","phone":"660793944","leadSource":"Walk into Store","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-22","status":"New","notes":"Waiting response for a meeting with owner","totalSales":0},{"id":101,"venue":"Southern Sun","channel":"B2B","firstName":"David","lastName":"Mokwebo","location":"","distributor":"","email":"david.mokwebo@southernsun.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"","totalSales":0},{"id":102,"venue":"Unclaimed","channel":"B2B","firstName":"Aidan","lastName":"","location":"Kloof Street","distributor":"","email":"","phone":"079 380 9210","leadSource":"Networking","accountManager":"Loydz","priority":"High","lastContacted":"2026-02-20","status":"Converted","notes":"","totalSales":0},{"id":103,"venue":"BMW Dealership","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"","totalSales":0},{"id":104,"venue":"Bidvest Lounge","channel":"B2B","firstName":"Suzanne","lastName":"VD","location":"","distributor":"","email":"SuzanneVD@bidvestcatering.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"Sent to procurement department","totalSales":0},{"id":105,"venue":"Molenvlliet","channel":"B2B","firstName":"Juliene","lastName":"","location":"","distributor":"","email":"stay@molenvliet.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":106,"venue":"Babylonstoren","channel":"B2B","firstName":"Elana","lastName":"","location":"","distributor":"","email":"functions@babylonstoren.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"High","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":107,"venue":"Boschendal","channel":"B2B","firstName":"Shanel","lastName":"","location":"","distributor":"","email":"functions@babylonstoren.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"High","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":108,"venue":"Da'Aria","channel":"B2B","firstName":"Danielle","lastName":"","location":"","distributor":"","email":"events@daria.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"High","lastContacted":"2026-04-28","status":"New","notes":"","totalSales":0},{"id":109,"venue":"Slow Lounge","channel":"B2B","firstName":"Janine","lastName":"","location":"","distributor":"","email":"janine@purefoods.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-23","status":"New","notes":"Waiting for response","totalSales":0},{"id":110,"venue":"Youngblood","channel":"B2B","firstName":"","lastName":"","location":"CBD","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":111,"venue":"OK Liquors Parow","channel":"Trade","firstName":"Denver","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"844552084","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"New","notes":"Will talk to owner and get back","totalSales":0},{"id":112,"venue":"OK Liquors Tygerdal","channel":"Trade","firstName":"Ray","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"211090099","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":113,"venue":"Tops Bellville","channel":"Trade","firstName":"Yushrah","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"638150626","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":114,"venue":"Tops Boston","channel":"Trade","firstName":"Valencia","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"653915424","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":115,"venue":"Tops Glenwood","channel":"Trade","firstName":"Obert","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"787540057","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":116,"venue":"Tops Parow Valley","channel":"Trade","firstName":"Fiona","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"746320000","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":117,"venue":"Observatory Bottle Store","channel":"Trade","firstName":"Dani","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"719445117","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-16","status":"New","notes":"Not interested yet","totalSales":0},{"id":118,"venue":"Blue Bottle Observatory","channel":"Trade","firstName":"Ling","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-16","status":"New","notes":"Owner will get back to me","totalSales":0},{"id":119,"venue":"Blue Bottle Panorama (Day 3)","channel":"Trade","firstName":"Carl","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"842054604","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":120,"venue":"Market Liquors Parklands Sandowne","channel":"Trade","firstName":"Adriana","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"822325842","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":121,"venue":"Tops Belhar","channel":"Trade","firstName":"Jonothan","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"673236926","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-13","status":"New","notes":"Sent owner a mail, no feedback","totalSales":0},{"id":122,"venue":"Blue Bottle Wild Liquors","channel":"Trade","firstName":"Evan","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"745206017","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":123,"venue":"Liquor City Paarl","channel":"Trade","firstName":"Zheng","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"766767666","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":124,"venue":"Market Liquors Malmesbury","channel":"Trade","firstName":"Janicka","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"840444024","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":125,"venue":"Market Liquors Paarl (Day 1)","channel":"Trade","firstName":"Joy","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"672330172","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":126,"venue":"OK Liquors Rembrandt Mall","channel":"Trade","firstName":"Burton","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"New","notes":"Under new ownership","totalSales":0},{"id":127,"venue":"Tops Zomerlust (Day 1)","channel":"Trade","firstName":"Dimitri","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"769262883","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":128,"venue":"Blue Bottle Liquors Paarl Mall","channel":"Trade","firstName":"Louise","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-04-02","status":"New","notes":"Not interested yet","totalSales":0},{"id":129,"venue":"Blue Bottle Liquor for Africa","channel":"Trade","firstName":"Yolande","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"836072111","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":130,"venue":"Market Liquors Brackenfell","channel":"Trade","firstName":"Elizabeth","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"612727071","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":131,"venue":"OK Liquors Goedemoed","channel":"Trade","firstName":"Dricus","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"719122379","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":132,"venue":"Tops Brackenfell","channel":"Trade","firstName":"Johan","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"652973469","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":133,"venue":"Tops Kraaifontein (Day 8)","channel":"Trade","firstName":"Adriaan","lastName":"","location":"Cape Town","distributor":"Karino","email":"kraaifonteinsuperspar@retail.co.za","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":134,"venue":"Tops Protea Hoogte (Day 8)","channel":"Trade","firstName":"LJ","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"712375157","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":135,"venue":"Tops Helshoogte","channel":"Trade","firstName":"Patuma","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"763409494","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":136,"venue":"Carolines Tokai","channel":"Trade","firstName":"Caroline","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"New","notes":"Not interested","totalSales":0},{"id":137,"venue":"Market Liquors Hermanus","channel":"Trade","firstName":"JJ","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"798977999","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":138,"venue":"OK Liquors Napier","channel":"Trade","firstName":"Ben","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"822576721","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"New","notes":"No feedback whatsoever","totalSales":0},{"id":139,"venue":"Tops Kleinmond (Day 9)","channel":"Trade","firstName":"Marjana","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"798331149","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":140,"venue":"Tops Villiersdorp (Day 9)","channel":"Trade","firstName":"Elaine","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"662807287","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":141,"venue":"Tops Welgevonden (Day 7)","channel":"Trade","firstName":"Wayne","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"621274008","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":142,"venue":"Market Liquors Tokai (Day 10)","channel":"Trade","firstName":"Kim","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"730617906","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":143,"venue":"OK Liquors Spineroad","channel":"Trade","firstName":"Eleanor","lastName":"","location":"Cape Town","distributor":"Karino","email":"spineroad@ok.co.za","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"New","notes":"Under new ownership","totalSales":0},{"id":144,"venue":"Tops Fishhoek","channel":"Trade","firstName":"Gilbert","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"768108940","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"Listed VS & VSOP","totalSales":0},{"id":145,"venue":"Tops Glencairn","channel":"Trade","firstName":"Theresa","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Cold Call","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":146,"venue":"Tops Houtbay","channel":"Trade","firstName":"Farai","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"710779622","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":147,"venue":"Tops Muizenberg","channel":"Trade","firstName":"Denise","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"820525529","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":148,"venue":"Tops Oakhurst","channel":"Trade","firstName":"Chicco","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"609232264","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":149,"venue":"Tops Weltevreden (Day 10)","channel":"Trade","firstName":"Paul","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"610052357","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":150,"venue":"Blue Bottle Elands Bay","channel":"Trade","firstName":"Patrick","lastName":"","location":"West-Coast","distributor":"Karino","email":"","phone":"","leadSource":"Website","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-23","status":"New","notes":"Will give feedback in April","totalSales":0},{"id":151,"venue":"Tops Edgemead","channel":"Trade","firstName":"Nolene","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"724245680","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-05","status":"New","notes":"Not interested yet","totalSales":0},{"id":152,"venue":"Tops Bredasdorp","channel":"Trade","firstName":"Anna","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"728264575","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Low","lastContacted":"2026-03-03","status":"Converted","notes":"","totalSales":0},{"id":153,"venue":"Tops Aurora (Day 6)","channel":"Trade","firstName":"Alcino","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"765381927","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":154,"venue":"Tops Durbanville","channel":"Trade","firstName":"Devon","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"815151023","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":155,"venue":"Tops Haasendal","channel":"Trade","firstName":"Tyrese","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"677984010","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":156,"venue":"Tops Old Oak","channel":"Trade","firstName":"Hennie","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"219192711","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":157,"venue":"Tops Palm Grove","channel":"Trade","firstName":"Charl","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"219762148","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":158,"venue":"Tops Phesantekraal","channel":"Trade","firstName":"Jestin","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"605484307","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-23","status":"Converted","notes":"","totalSales":0},{"id":159,"venue":"Tops Kuilsriver (Day 5)","channel":"Trade","firstName":"Kevin","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"785021144","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":160,"venue":"Tops Sunrise Circle (Day 5)","channel":"Trade","firstName":"Barry","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"738878330","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-21","status":"Converted","notes":"","totalSales":0},{"id":161,"venue":"Tops Mowbray","channel":"Trade","firstName":"Michelle","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"765520284","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":162,"venue":"Tops Observatory","channel":"Trade","firstName":"Lincon","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"658430389","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":163,"venue":"Tops Wex1","channel":"Trade","firstName":"Ernst","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"738837072","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":164,"venue":"Tops Woodstock Quarter","channel":"Trade","firstName":"Ernst","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"738837072","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":165,"venue":"Whisky Shop","channel":"Trade","firstName":"Hector","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"833771113","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":166,"venue":"Liquor City Long Street (Day 4)","channel":"Trade","firstName":"Arthur","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"680530467","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":167,"venue":"Woodstock Liquors","channel":"Trade","firstName":"Inacio","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"832689430","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":168,"venue":"Tops Adderly","channel":"Trade","firstName":"Noni","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":169,"venue":"Tops Cape Station","channel":"Trade","firstName":"Noni","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":170,"venue":"Tops Century Village","channel":"Trade","firstName":"Anthony","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"825215452","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-16","status":"Converted","notes":"","totalSales":0},{"id":171,"venue":"Blouberg Liquors","channel":"Trade","firstName":"Inacio","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"832689430","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":172,"venue":"NGF Sunset Beach","channel":"Trade","firstName":"Gary","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"726145099","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":173,"venue":"OK Liquors Duynefontein","channel":"Trade","firstName":"Frean","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"722437068","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":174,"venue":"Tops Blouberg","channel":"Trade","firstName":"Canais","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"635074577","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":175,"venue":"Tops Melkbos","channel":"Trade","firstName":"Abigail","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"715768900","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":176,"venue":"Tops Milnerton","channel":"Trade","firstName":"Cheslyn","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"692649715","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":177,"venue":"Tops Morningfield (Day 3)","channel":"Trade","firstName":"Sinetemba","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"833584184","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":178,"venue":"Tops Parklands","channel":"Trade","firstName":"Morne","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"743749296","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":179,"venue":"Tops Riberios","channel":"Trade","firstName":"Louise","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"673611248","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":180,"venue":"Tops Royal Ascot","channel":"Trade","firstName":"Tumeka","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"632692784","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":181,"venue":"Tops Sunningdale","channel":"Trade","firstName":"Tinotenda","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"658334199","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":182,"venue":"Tops Tableview","channel":"Trade","firstName":"Roger","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"718921187","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-14","status":"Converted","notes":"","totalSales":0},{"id":183,"venue":"Tops Malmesbury","channel":"Trade","firstName":"Hendika","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"785751687","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":184,"venue":"Tops Vineyard","channel":"Trade","firstName":"Christiaan","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"720407583","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-09","status":"Converted","notes":"","totalSales":0},{"id":185,"venue":"Tops Cape Quarter (Day 2)","channel":"Trade","firstName":"Chantal","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"836613653","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":186,"venue":"Barkeeper","channel":"Trade","firstName":"Carla","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"823391079","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":187,"venue":"Constantia Wine & Craft","channel":"Trade","firstName":"Jeremy","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"786424044","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":188,"venue":"NGF Gardens","channel":"Trade","firstName":"Vuyo","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"7433425032","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":189,"venue":"NGF Three Anchor Bay","channel":"Trade","firstName":"Joline","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"786584312","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":190,"venue":"Tops Bergvliet","channel":"Trade","firstName":"Vuyi","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"766611750","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"New","notes":"Not interested yet","totalSales":0},{"id":191,"venue":"Tops Kloof","channel":"Trade","firstName":"Makhumalo","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"631053550","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":192,"venue":"Tops Rosmead (Day 2)","channel":"Trade","firstName":"Rochelle","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"728603028","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":193,"venue":"Tops Seapoint","channel":"Trade","firstName":"Stuart","lastName":"","location":"CBD","distributor":"Karino","email":"seapoint1@retail.spar.co.za","phone":"","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":194,"venue":"Tops Vredehoek","channel":"Trade","firstName":"Lisa","lastName":"","location":"CBD","distributor":"Karino","email":"","phone":"682316301","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-04-07","status":"Converted","notes":"","totalSales":0},{"id":195,"venue":"Tops Bottelary (Day 7)","channel":"Trade","firstName":"Jestin","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"605484307","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-27","status":"Converted","notes":"","totalSales":0},{"id":196,"venue":"OK Liquors Urban Sonstraal","channel":"Trade","firstName":"Bri-Niel","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"733719691","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":197,"venue":"Tops Cape Gate","channel":"Trade","firstName":"Frank","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"638538885","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":198,"venue":"Liquor City Somerset Mall","channel":"Trade","firstName":"Charlene","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"715243228","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":199,"venue":"OK Liquors Strand","channel":"Trade","firstName":"Meaghan","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"733695375","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":200,"venue":"Tops Cinnamon","channel":"Trade","firstName":"Tafadswa","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"614817661","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":201,"venue":"Tops De Jonker","channel":"Trade","firstName":"Charlton","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"641629569","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":202,"venue":"Tops Die Boord","channel":"Trade","firstName":"Gershwin","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"613656818","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":203,"venue":"Tops Helderbosch","channel":"Trade","firstName":"Chantal","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"674027744","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":204,"venue":"Tops Helderveu","channel":"Trade","firstName":"Heinrich","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"619525640","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":205,"venue":"Tops Lions Square","channel":"Trade","firstName":"Barbara","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"790973527","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":206,"venue":"Tops Mountainview","channel":"Trade","firstName":"Melisia","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"740583958","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":207,"venue":"Tops Paradyskloof","channel":"Trade","firstName":"Kenny","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"739506767","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":208,"venue":"Tops Paul Roos","channel":"Trade","firstName":"Nathely","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"780743365","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":209,"venue":"Tops Strand","channel":"Trade","firstName":"Lycken","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"620494468","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"New","notes":"Not interesteed yet","totalSales":0},{"id":210,"venue":"Tops Twin Palms","channel":"Trade","firstName":"Senobia","lastName":"","location":"Cape Winelands","distributor":"Karino","email":"","phone":"843573643","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-26","status":"Converted","notes":"","totalSales":0},{"id":211,"venue":"Liquor City Claremont","channel":"Trade","firstName":"Derek","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"216741478","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":212,"venue":"OK Liquors Ridgeworth","channel":"Trade","firstName":"Quinton","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"837799296","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":213,"venue":"Picardi Rebel Claremont (Day 4)","channel":"Trade","firstName":"Tristian","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"662086625","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":214,"venue":"Tops Eastcliff","channel":"Trade","firstName":"Nadia","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"790927628","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"Converted","notes":"","totalSales":0},{"id":215,"venue":"Tops Eersteriver (Day 6)","channel":"Trade","firstName":"Sinalo","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"735920264","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-25","status":"New","notes":"Not interested at all","totalSales":0},{"id":216,"venue":"OK Liquors Strandfontein","channel":"Trade","firstName":"Jenique","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"726130092","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":217,"venue":"Tops Alphen","channel":"Trade","firstName":"Riyaaz","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"742624147","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-24","status":"Converted","notes":"","totalSales":0},{"id":218,"venue":"Ultra Nyanga","channel":"Trade","firstName":"Sisa","lastName":"Liwani","location":"Nyanga","distributor":"","email":"nyanga@ultraliquors.co.za","phone":"082 398 0404","leadSource":"Online","accountManager":"Loydz","priority":"Medium","lastContacted":"2026-02-18","status":"New","notes":"","totalSales":0},{"id":219,"venue":"Tops Sonstraal","channel":"Trade","firstName":"Selvinia","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"634486222","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":220,"venue":"Tops Vredekloof","channel":"Trade","firstName":"Riaan","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"605905596","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":221,"venue":"Tops Welgelegeen","channel":"Trade","firstName":"Yianakis","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"720615896","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":222,"venue":"Tops Zevenwacht","channel":"Trade","firstName":"Wesley","lastName":"","location":"Cape Town","distributor":"Karino","email":"","phone":"824777793","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-19","status":"Converted","notes":"","totalSales":0},{"id":223,"venue":"Dal Italia Restaurant","channel":"Trade","firstName":"Fabio","lastName":"","location":"Overberg","distributor":"Karino","email":"","phone":"823372339","leadSource":"Call-cycle","accountManager":"Lehmarc","priority":"Medium","lastContacted":"2026-03-06","status":"Converted","notes":"","totalSales":0},{"id":224,"venue":"City Lodge Hotels","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":225,"venue":"Anew Hotels","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":226,"venue":"Bon Hotels","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":227,"venue":"Steyn City","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":228,"venue":"Taj Hotel","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":229,"venue":"Glendower","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":230,"venue":"Randpark Golf Club","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":231,"venue":"Houghton Golf Club","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-04-01","status":"New","notes":"","totalSales":0},{"id":232,"venue":"Londolozi","channel":"B2B","firstName":"Duncan","lastName":"","location":"","distributor":"","email":"duncan@londolozi.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"No response, call today","totalSales":0},{"id":233,"venue":"MalaMala","channel":"B2B","firstName":"Rufie or Ross","lastName":"","location":"","distributor":"","email":"","phone":"781008810","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"No response, call today","totalSales":0},{"id":234,"venue":"Shamwari Group","channel":"B2B","firstName":"Janine","lastName":"","location":"","distributor":"","email":"procurement@shamwari.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"","totalSales":0},{"id":235,"venue":"Singita (Group)","channel":"B2B","firstName":"Schwess","lastName":"","location":"","distributor":"","email":"capepremierwine@singita.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"2026-03-24","status":"New","notes":"","totalSales":0},{"id":236,"venue":"Foodbarn Group","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":237,"venue":"Hussar Grill Group","channel":"B2B","firstName":"Bradley","lastName":"","location":"","distributor":"","email":"bradleyh@spurcorp.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":238,"venue":"Newmark Hotels","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":239,"venue":"Burkenhead House (RP)","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":240,"venue":"Ellerman House","channel":"B2B","firstName":"Manuel","lastName":"","location":"","distributor":"","email":"manuel@ellerman.co.za","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":241,"venue":"Fairmont / Accor","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":242,"venue":"Hilton SA","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":243,"venue":"Kapama Private Game Reserve","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":244,"venue":"La Residence (RP)","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":245,"venue":"One&Only Cape Town","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":246,"venue":"Royal Malewane (RP)","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":247,"venue":"Sun International","channel":"B2B","firstName":"Mpho","lastName":"Moshotle","location":"","distributor":"","email":"Mpho.moshotle@suninternational.com","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":248,"venue":"Ulusaba (Virgin)","channel":"B2B","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Alex","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":249,"venue":"Energy Masterbuilders Paarl","channel":"Trade","firstName":"Johan","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":250,"venue":"Hoekstra Farming","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":251,"venue":"Nexus Agriculture","channel":"Trade","firstName":"Ester","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":252,"venue":"PSG Paarl","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":253,"venue":"PSG Stellenbosch","channel":"Trade","firstName":"Joaan","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":254,"venue":"BMW Paarlberg","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":255,"venue":"BMW Stellenbosch","channel":"Trade","firstName":"Christopher","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":256,"venue":"GWM Paarl","channel":"Trade","firstName":"Brandon","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":257,"venue":"Exsa","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":258,"venue":"Paarl Boys High Waterpolo","channel":"Trade","firstName":"Ross","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":259,"venue":"Stellenberg High School","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":260,"venue":"New Orleans High School","channel":"Trade","firstName":"Carlene","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":261,"venue":"Nuweland Paarl","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":262,"venue":"Giflo Group","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":263,"venue":"Samsung","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":264,"venue":"Icon Fruit","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":265,"venue":"Isipani Paarl","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":266,"venue":"Sky vines","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":267,"venue":"Akura Paarl","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":268,"venue":"Rawson H/Q","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":269,"venue":"Workshop 17","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":270,"venue":"Cargokwik","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":271,"venue":"Zuus Paarl","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":272,"venue":"Selavie (Val de Vie)","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":273,"venue":"Blacksmith Kitchen","channel":"Trade","firstName":"Bianca","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":274,"venue":"Grande Roche Hotel","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":275,"venue":"Noop","channel":"Trade","firstName":"Zian","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":276,"venue":"Laborie Estate","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":277,"venue":"Veld and Vine","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0},{"id":278,"venue":"Val de Vie","channel":"Trade","firstName":"","lastName":"","location":"","distributor":"","email":"","phone":"","leadSource":"Cold Call","accountManager":"Unassigned","priority":"Medium","lastContacted":"","status":"New","notes":"","totalSales":0}];
const LOGO_URL = "data:image/png;base64,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nHPOOeecc84555xzzjnnnHPOOeecc84555xzzjnnnHPOOeecc84555xzzjnnnHPOOeecc84555xzzjnnnHPOOeecc84555xzzjnnnHPOOeecc84555xzzjnnnHPOOeecc84555xzzjnnnHPOOeecc84555xzzjnnnHPOOeecc84555xzzjnnnHPOOeec2ybI9vAmVk6+3VZPuBUkkkhEgyQQxTCTN/gARjSUhBQxIIhR9NiRnod/6g37h5c+cImtWz4NDUbQSAhGSSBqIEQoSYhEVCBI3PiNoXpJIX9OJP+3GCARYkDUUFFECkwSrboOYIc937dNtJNFk6+21StnEi1Qo5GYEoqBKgFFtcTMMEugipEwFcQSWMK0xCwhpvm/zRBTSlWiKWoJpAQVdj/tn2/YMZk1+jJbMPoKkoAomBkASUL+AjMwya9P8v/nrwj5r8yavweg+U9m1dcYpvlnqQAoYgFVg5C/TCx/pxl07rsHB73vN9tUXzHy75+xpQvGVc2+QKTEgMLAJBKDAoaKIAghSHWkDBFDms6P0HSeKyJCBCRUR1uE1p0GsMfpf9mqjs3S8dfbqknXkyQgUrXv0BpRw4IRtA4JChY368dyfxAwEaTqN8wSEgOoYAJCaP7eDrsdTaedT3hF733FlAds9ZS7sRib210+zrKxT7KASlP/pECAIM39VX59svnlq6lvq97Dxv4t/9mCgYXmr93075r+LRFpPldEAiaAVufYJl+j1bmW/2ybfD5sPJ7VeflaheYXVNK+x7507L7HKzreq5ZMsDULx5NC0zmdf2YKUKRACtqiA1A0pW2Pvei0w56v+RwutocDUls5nw3zc4OJKbIuKHVmNGyBrChVYMunvKKhoB32hv6b65dOoWHRU5RiRAmUCFGUQP6zxAKSIDFs1hGANHdETZ9TqTrSoCh1RIwkRowRrCCG1ttMO1izag4bVkxDkqIkTGugCUslVAFJrYRkYDlYoRvy5zVhVYAilVWYUqQsURpRFSAhqYSU3tD3US6fzuLpDxJDQKsOualfNtvY4dsm/aBavlgYOUzpJpfJpp8hmi8CqQpVOUhrvugmy21BDSRQy1etKmjbNtcnLF04gUVTRlEXSswCGhORmM9XUSQYhchm5zEiRJTQdJPRfO01pPlarUh14Q4iBHmAmY/83AYe/KWtJlBtWDWbNfPGoAp1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"

const SALES_REPS = ['Alex', 'Lehmarc', 'Loydz', 'Louis', 'Matthew', 'Anthony'];

// Sales rep email addresses — used in the signature of order confirmation emails
// so the client knows who the message is from and where to reply.
const REP_EMAILS = {
  Alex: 'alex@breakfreebeverages.com',
  Lehmarc: 'lehmarc@breakfreebeverages.com',
  Loydz: 'lonwabo@breakfreebeverages.com',
  Louis: 'louis@breakfreebeverages.com',
  Matthew: 'matthew@breakfreebeverages.com',
  Anthony: 'anthony@breakfreebeverages.com',
};

// Personal/internal addresses that must NEVER be auto-populated as a
// recipient anywhere in the app (test data leftovers etc.)
const BLOCKED_EMAILS = ['mrunnalls22@gmail.com'];
const isBlockedEmail = (email) => !!email && BLOCKED_EMAILS.includes(String(email).trim().toLowerCase());

// Build a mailto: link that notifies a tagged agent
function composeTagEmail({ rep, clientName, location, channel, taggedBy, type, date, outcome, notes, followUp, status }) {
  const to = REP_EMAILS[rep] || '';
  const subject = `You've been tagged on ${clientName} — Avante CRM`;
  const lines = [];
  lines.push(`Hi ${rep},`);
  lines.push('');
  lines.push(`You've been tagged on a client in the Avante Sales CRM.`);
  lines.push('');
  lines.push(`CLIENT:    ${clientName}`);
  if (location) lines.push(`LOCATION:  ${location}`);
  if (channel) lines.push(`CHANNEL:   ${channel}`);
  if (taggedBy) lines.push(`TAGGED BY: ${taggedBy}`);
  lines.push(`TYPE:      ${type === 'visit' ? 'Visit Note' : 'Client Tag'}`);
  lines.push('');
  if (type === 'visit') {
    if (date) lines.push(`VISIT DATE: ${date}`);
    if (outcome) lines.push(`OUTCOME:    ${outcome}`);
    if (notes) lines.push(`NOTES:      "${notes}"`);
    if (followUp) lines.push(`FOLLOW-UP:  ${followUp}`);
  } else {
    if (status) lines.push(`STATUS:     ${status}`);
    if (taggedBy) lines.push(`ACCOUNT MANAGER: ${taggedBy}`);
    if (notes) lines.push(`NOTES:      "${notes}"`);
  }
  lines.push('');
  lines.push('View this client in the CRM:');
  lines.push('https://avante-crm.vercel.app/ \u2192 Notifications');
  lines.push('');
  lines.push('\u2014 Avante Sales CRM');
  lines.push('   Dare to Forward');

  const body = lines.join('\n');
  return `mailto:${to}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}

// Open a mailto link for each newly-tagged rep (only those not already in prevTags)
function notifyNewTags({ prevTags = [], newTags = [], ...rest }) {
  const added = newTags.filter(r => !prevTags.includes(r));
  added.forEach((rep, i) => {
    const url = composeTagEmail({ rep, ...rest });
    const fire = () => {
      const a = document.createElement('a');
      a.href = url;
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    };
    // First link fires immediately (within the click gesture).
    // Any additional tagged reps fire shortly after.
    if (i === 0) fire();
    else setTimeout(fire, i * 500);
  });
  return added;
}

const CHANNELS = ['Trade', 'On-Con', 'B2B'];
const PAYMENT_TERMS = ['COD', '30 Days', '60 Days'];
const CONTACT_METHODS = ['In Person', 'WhatsApp', 'Phone Call / Online Meet', 'Email'];
const LOCATIONS = [
  'Western Cape', 'Eastern Cape', 'Northern Cape', 'Gauteng',
  'KwaZulu-Natal', 'Free State', 'Limpopo', 'Mpumalanga', 'North West', 'Other',
];
const STATUSES = ['New', 'Contacted', 'Converted', 'Lost'];

// B2B-specific pipeline stages with probability percentages
const B2B_STATUSES = [
  { label: 'New',                  pct: 1   },
  { label: 'Contacted',            pct: 10  },
  { label: 'Discovery Completed',  pct: 40  },
  { label: 'Pitched',              pct: 60  },
  { label: 'Converted / Signed',   pct: 80  },
  { label: 'Invoiced / Paid',      pct: 100 },
];
const B2B_STATUS_LABELS = B2B_STATUSES.map(s => s.label);

// Returns the right status list for a given channel
const getStatusesForChannel = (channel) =>
  channel === 'B2B' ? B2B_STATUS_LABELS : STATUSES;

// Returns the probability % for a B2B status label, or null for non-B2B
const getB2BPct = (statusLabel) => {
  const found = B2B_STATUSES.find(s => s.label === statusLabel);
  return found ? found.pct : null;
};

// A visit counts as a "sale" for conversion-rate purposes if either:
// 1. The outcome was explicitly logged as 'Sold In', or
// 2. SKU items with quantity were placed (a real order), or
// 3. It has a recorded saleAmount > 0
// This keeps conversion rate consistent with the auto-convert-to-Converted logic in addVisit.
const isSoldVisit = (v) =>
  v.outcome === 'Sold In' ||
  (Number(v.saleAmount) > 0) ||
  (Array.isArray(v.items) && v.items.some(it => Number(it.qty) > 0));

// Unified status colour lookup (standard + B2B)
const getStatusColor = (status) => ({
  'New':                 '#5A7A99',
  'Contacted':           '#BC8D26',
  'Converted':           '#2d8659',
  'Lost':                '#CC233A',
  'Discovery Completed': '#5A7A99',
  'Pitched':             '#BC8D26',
  'Converted / Signed':  '#2d8659',
  'Invoiced / Paid':     '#1a5c38',
}[status] || '#5A7A99');
const LEAD_SOURCES = ['Cold Call', 'Referral', 'Walk into Store', 'Email', 'Event', 'Networking', 'Call-cycle', 'Website', 'Online'];
const PRIORITIES = ['Low', 'Medium', 'High'];

const DEFAULT_TARGETS = {
  Alex:    { revenue: 150000, visits: 60, tradeRetail: 20, onCon: 0, b2b: 40 },
  Lehmarc: { revenue: 150000, visits: 60, tradeRetail: 40, onCon: 0, b2b: 20 },
  Loydz:   { revenue: 100000, visits: 45, tradeRetail: 15, onCon: 0, b2b: 30 },
  Louis:   { revenue: 100000, visits: 45, tradeRetail: 15, onCon: 0, b2b: 30 },
  Matthew: { revenue: 100000, visits: 45, tradeRetail: 15, onCon: 0, b2b: 30 },
  Anthony: { revenue: 100000, visits: 45, tradeRetail: 15, onCon: 0, b2b: 30 },
};

// Avante SKU catalogue (Trade Ex VAT prices, ZAR per unit)
const SKU_CATALOGUE = [
  { id: 'vsop_750_current', name: 'VSOP 750ml (Current)', price: 420.00 },
  { id: 'vsop_750_new',     name: 'VSOP 750ml (New)',     price: 579.00 },
  { id: 'vsop_200',         name: 'VSOP 200ml',           price: 150.00 },
  { id: 'xv_750',           name: 'XV 750ml',             price: 2086.00 },
  { id: 'vs_500',           name: 'VS 500ml',             price: 217.00 },
  { id: 'vs_750',           name: 'VS 750ml',             price: 296.50 },
  { id: 'xv_200',           name: 'XV 200ml',             price: 745.00 },
  { id: 'xo_750',           name: 'XO 750ml',             price: 1043.00 },
  { id: 'vs_200',           name: 'VS 200ml',             price: 113.00 },
  { id: 'gift_4x50',        name: '4 x 50ml Gift Box',    price: 520.00 },
  { id: 'gift_3x200',       name: '3 x 200ml Gift Box',   price: 1050.00 },
  { id: 'custom_vs',        name: 'Custom VS',            price: 694.78 },
  { id: 'custom_xv',        name: 'Custom XV',            price: 694.78 },
  // 50ml singles & mixed box
  { id: 'mixed_50ml_box',   name: 'Mixed 50ml Box',       price: 961.50 },
  { id: 'vs_50ml',          name: 'VS 50ml',              price: 654.00 },
  { id: 'vsop_50ml',        name: 'VSOP 50ml',            price: 768.00 },
  { id: 'xo_50ml',          name: 'XO 50ml',              price: 1080.00 },
  { id: 'xv_50ml',          name: 'XV 50ml',              price: 1344.00 },
];

// Get effective SKU catalogue with any price overrides applied
const getEffectiveSkus = (overrides = {}) =>
  SKU_CATALOGUE.map(s => ({ ...s, price: overrides[s.id] ?? s.price }));

const ZAR = (n) => 'R ' + Math.round(n || 0).toLocaleString('en-ZA');
const todayISO = () => new Date().toISOString().slice(0, 10);
const monthISO = () => new Date().toISOString().slice(0, 7);

// =================== Avante Logo ===================
// Logo is landscape 595x417 — use height + auto width
// AvanteLogo — transparent PNG with gold artwork, no white box
const AvanteLogo = ({ className = '', height = 52, gold = true }) => (
  <img
    src={LOGO_URL}
    alt="Avante Cape Brandy"
    className={className}
    style={{
      height: height,
      width: 'auto',
      objectFit: 'contain',
      display: 'block',
      maxWidth: height * 1.8,
      // On dark header, logo is already gold. On light bg, tint darker.
      filter: gold ? 'none' : 'brightness(0.3) sepia(1) hue-rotate(0deg)',
    }}
    onError={e => { e.target.style.display = 'none'; }}
  />
);

const AvanteWordmark = ({ dark = false, className = '' }) => (
  <svg viewBox="0 0 280 60" className={className} xmlns="http://www.w3.org/2000/svg">
    <text x="140" y="38" textAnchor="middle" fontFamily="'Cinzel', 'Copperplate', serif" fontWeight="700" fontSize="34" letterSpacing="6" fill={dark ? '#FCF7F2' : '#002855'}>AVANTE</text>
    <text x="140" y="55" textAnchor="middle" fontFamily="'Cinzel', 'Copperplate', serif" fontWeight="600" fontSize="9" letterSpacing="6" fill={dark ? '#FCF7F2' : '#002855'}>CAPE BRANDY</text>
  </svg>
);

const RaysBackdrop = ({ opacity = 0.07 }) => (
  <svg viewBox="0 0 400 400" className="absolute inset-0 w-full h-full pointer-events-none" preserveAspectRatio="xMidYMid slice" xmlns="http://www.w3.org/2000/svg">
    <g stroke="#DBB85E" strokeWidth="0.4" opacity={opacity}>
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
  const obj = {
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
    payment_terms: c.paymentTerms || '',
    client_tags: c.clientTags || [],
    prospected_amount: c.prospectedAmount || 0,
  };
  if (c.id !== undefined) obj.id = c.id;
  return obj;
}

function clientFromDb(r) {
  return {
    id: Number(r.id),
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
    paymentTerms: r.payment_terms || '',
    clientTags: r.client_tags || [],
    prospectedAmount: Number(r.prospected_amount) || 0,
  };
}

// Columns confirmed to exist in the visits table (base schema)
// Run fix_visits_columns.sql to add: channel, contact_method, follow_up_date, tagged_reps
const VISITS_EXTENDED_COLS = { channel: true, contact_method: true, follow_up_date: true, tagged_reps: true };

function visitToDb(v, includeExtended = true) {
  const row = {
    client_id: Number(v.clientId),
    sales_rep: v.salesRep || '',
    date: v.date || null,
    outcome: v.outcome || '',
    sale_amount: v.saleAmount || 0,
    items: v.items || [],
    notes: v.notes || '',
    follow_up_notes: v.followUpNotes || v.followUp || '',
    order_placed: v.orderPlaced || '',
  };
  if (includeExtended) {
    row.contact_method = v.contactMethod || '';
    row.follow_up_date = v.followUpDate || null;
    row.tagged_reps = v.taggedReps || [];
    row.channel = v.channel || '';
  }
  return row;
}

function visitFromDb(r) {
  return {
    id: Number(r.id),         // always a number for consistent equality checks
    clientId: Number(r.client_id),
    salesRep: r.sales_rep || '',
    date: r.date || '',
    outcome: r.outcome || '',
    saleAmount: Number(r.sale_amount) || 0,
    items: r.items || [],
    notes: r.notes || '',
    followUpNotes: r.follow_up_notes || '',
    orderPlaced: r.order_placed || '',
    contactMethod: r.contact_method || '',
    followUpDate: r.follow_up_date || '',
    taggedReps: r.tagged_reps || [],
    channel: r.channel || '',
  };
}

function prospectToDb(p) {
  return {
    sales_rep: p.salesRep || '',
    client_name: p.clientName || '',
    sku_notes: p.skuNotes || '',
    amount: Number(p.amount) || 0,
    expected_date: p.expectedDate || null,
    status: p.status || 'Open',
    client_id: p.clientId ? Number(p.clientId) : null,
    items: p.items || [],
  };
}

function prospectFromDb(r) {
  return {
    id: r.id,
    salesRep: r.sales_rep || '',
    clientName: r.client_name || '',
    clientId: r.client_id || null,
    skuNotes: r.sku_notes || '',
    amount: Number(r.amount) || 0,
    expectedDate: r.expected_date || '',
    status: r.status || 'Open',
    items: r.items || [],
    createdAt: r.created_at || '',
  };
}

// ── B2B CUSTOMS — production tracker (standalone, not linked to clients table) ──
function b2bCustomToDb(row) {
  return {
    customer_name:    row.customerName || '',
    deposit_paid:     row.depositPaid || 'No',
    briefed:          row.briefed || 'No',
    liquid_lined_up:  row.liquidLinedUp || 'No',
    production_stage: row.productionStage || 'All goods in warehouse',
    balance_paid:     row.balancePaid || 'No',
    ready_dispatch:   row.readyDispatch || 'No',
  };
}

function b2bCustomFromDb(r) {
  return {
    id: Number(r.id),
    customerName:    r.customer_name || '',
    depositPaid:     r.deposit_paid || 'No',
    briefed:         r.briefed || 'No',
    liquidLinedUp:   r.liquid_lined_up || 'No',
    productionStage: r.production_stage || 'All goods in warehouse',
    balancePaid:     r.balance_paid || 'No',
    readyDispatch:   r.ready_dispatch || 'No',
    createdAt:       r.created_at || '',
  };
}

function targetToDb(rep, t) {
  return {
    rep,
    revenue: t.revenue || 0,
    visits: t.visits || 0,
    trade_retail: t.tradeRetail || 0,
    on_con: t.onCon || 0,
    b2b: t.b2b || 0,
  };
}

function targetFromDb(r) {
  return {
    revenue: Number(r.revenue) || 0,
    visits: Number(r.visits) || 0,
    tradeRetail: Number(r.trade_retail) || 0,
    onCon: Number(r.on_con) || 0,
    b2b: Number(r.b2b) || 0,
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
        <div style={{ padding: 40, fontFamily: 'Georgia, serif', color: '#002855', textAlign: 'center', minHeight: '100vh', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', background: '#FCF7F2' }}>
          <div style={{ fontFamily: 'Cinzel, serif', fontSize: 20, fontWeight: 700, letterSpacing: '0.15em', color: '#BC8D26' }}>AVANTE CRM</div>
          <p style={{ marginTop: 16, color: '#002855' }}>Something went wrong. Please reload.</p>
          <pre style={{ marginTop: 12, fontSize: 11, color: '#CC233A', background: '#f9f5e8', padding: 12, maxWidth: 500, overflow: 'auto', textAlign: 'left' }}>{this.state.error?.message || String(this.state.error)}</pre>
          <button onClick={() => window.location.reload()} style={{ marginTop: 16, padding: '10px 24px', background: '#BC8D26', color: '#fff', border: 'none', cursor: 'pointer', fontFamily: 'Cinzel, serif', letterSpacing: '0.15em', fontSize: 11 }}>RELOAD</button>
        </div>
      );
    }
    return this.props.children;
  }
}


function LoginScreen({ onLogin }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = (e) => {
    e.preventDefault();
    setLoading(true);
    setError('');
    setTimeout(() => {
      const userConfig = USERS[email.trim().toLowerCase()];
      if (userConfig && userConfig.password === password) {
        const user = { email: email.trim().toLowerCase(), rep: userConfig.rep, role: userConfig.role, name: userConfig.name };
        storeUser(user);
        onLogin(user);
      } else {
        setError('Incorrect email or password. Please try again.');
      }
      setLoading(false);
    }, 400);
  };

  return (
    <div style={{ minHeight: '100vh', background: '#002855', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 24, fontFamily: "'Libre Baskerville', Georgia, serif", animation: 'fadeIn 0.3s ease-in' }}>
      {/* Logo area */}
      <div style={{ marginBottom: 40, textAlign: 'center' }}>
        <AvanteLogo height={72} />
        <p style={{ fontFamily: "'Cinzel', serif", fontSize: 10, letterSpacing: '0.5em', color: '#DBB85E', fontWeight: 600, marginTop: 16 }}>SALES CRM</p>
        <p style={{ fontFamily: "'Cinzel', serif", fontSize: 8, letterSpacing: '0.3em', color: 'rgba(255,255,255,0.4)', marginTop: 4 }}>FORWARD TOGETHER</p>
      </div>

      {/* Login card */}
      <div style={{ background: '#FCF7F2', width: '100%', maxWidth: 400, padding: '36px 32px', border: '2px solid rgba(188,141,38,0.4)' }}>
        <h2 style={{ fontFamily: "'Cinzel', serif", fontSize: 16, letterSpacing: '0.2em', color: '#002855', fontWeight: 700, textAlign: 'center', marginBottom: 8 }}>SIGN IN</h2>
        <p style={{ fontSize: 12, color: '#5A7A99', textAlign: 'center', fontStyle: 'italic', marginBottom: 28 }}>Sign in with your Break Free Beverages email</p>

        <form onSubmit={handleSubmit} autoComplete="on">
          <div style={{ marginBottom: 16 }}>
            <label style={{ fontFamily: "'Cinzel', serif", fontSize: 9, letterSpacing: '0.3em', color: '#BC8D26', fontWeight: 600, display: 'block', marginBottom: 6 }}>EMAIL ADDRESS</label>
            <input
              type="email"
              name="email"
              autoComplete="email"
              value={email}
              onChange={e => { setEmail(e.target.value); setError(''); }}
              placeholder="you@breakfreebeverages.com"
              required
              style={{ width: '100%', padding: '12px 14px', border: '1px solid rgba(0,40,85,0.2)', background: '#fff', fontFamily: "'Libre Baskerville', Georgia, serif", fontSize: 13, color: '#002855', outline: 'none', boxSizing: 'border-box' }}
            />
          </div>
          <div style={{ marginBottom: 24 }}>
            <label style={{ fontFamily: "'Cinzel', serif", fontSize: 9, letterSpacing: '0.3em', color: '#BC8D26', fontWeight: 600, display: 'block', marginBottom: 6 }}>PASSWORD</label>
            <input
              type="password"
              name="password"
              autoComplete="current-password"
              value={password}
              onChange={e => { setPassword(e.target.value); setError(''); }}
              placeholder="Enter your password"
              required
              style={{ width: '100%', padding: '12px 14px', border: '1px solid rgba(0,40,85,0.2)', background: '#fff', fontFamily: "'Libre Baskerville', Georgia, serif", fontSize: 13, color: '#002855', outline: 'none', boxSizing: 'border-box' }}
            />
          </div>
          {error && (
            <p style={{ fontSize: 11, color: '#CC233A', fontStyle: 'italic', marginBottom: 16, textAlign: 'center' }}>{error}</p>
          )}
          <button
            type="submit"
            disabled={loading}
            style={{ width: '100%', padding: '14px', background: loading ? 'rgba(0,40,85,0.4)' : '#002855', color: '#FCF7F2', border: 'none', fontFamily: "'Cinzel', serif", fontSize: 11, letterSpacing: '0.3em', fontWeight: 700, cursor: loading ? 'default' : 'pointer', transition: 'background 0.2s' }}>
            {loading ? 'SIGNING IN...' : 'SIGN IN'}
          </button>
        </form>
      </div>
      <p style={{ fontSize: 10, color: 'rgba(255,255,255,0.25)', marginTop: 24, fontStyle: 'italic' }}>Avante Cape Brandy · Confidential Internal Tool</p>
    </div>
  );
}

// ── AUTH WRAPPER ─────────────────────────────────────────────────────────────
// Keeps auth state separate from the app so login/logout doesn't remount
// the entire component tree and cause a white flash.
export default function AvanteCRM() {
  const [currentUser, setCurrentUser] = useState(() => getStoredUser());

  const handleLogin = (user) => {
    storeUser(user);
    setCurrentUser(user);
  };

  const handleLogout = () => {
    clearUser();
    setCurrentUser(null);
  };

  if (!currentUser) {
    return <LoginScreen onLogin={handleLogin} />;
  }

  return <AvanteCRMApp currentUser={currentUser} onLogout={handleLogout} />;
}

// ── MAIN APP ──────────────────────────────────────────────────────────────────
function AvanteCRMApp({ currentUser, onLogout }) {

  // Helper shortcuts
  const userRep = currentUser.rep;       // e.g. 'Alex'
  const userIsManager = isManager(currentUser);

  // ── APP STATE ─────────────────────────────────────────────────────────────
  const [view, setView] = useState('dashboard');
  const [loading, setLoading] = useState(true);
  const [clients, setClients] = useState([]);
  const [visits, setVisits] = useState([]);
  const [targets, setTargets] = useState(DEFAULT_TARGETS);
  const [skuPrices, setSkuPrices] = useState({});
  const [prospects, setProspects] = useState([]);  // overrides for SKU prices set in Manager Portal
  const [b2bCustoms, setB2bCustoms] = useState([]); // B2B Customs production tracker rows
  const [managerAuthed, setManagerAuthed] = useState(false); // persists across tab changes
  const [placeOrderClient, setPlaceOrderClient] = useState(null); // client to place order for
  // Reps default to their own data; manager defaults to 'All'
  const [activeRep, setActiveRep] = useState(userIsManager ? 'All' : userRep);
  const [activeMonth, setActiveMonth] = useState(monthISO()); // e.g. "2026-06"
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

  // Non-managers only see clients assigned to them
  const visibleClients = useMemo(() =>
    userIsManager ? clients : clients.filter(c => c.accountManager === userRep),
    [clients, userIsManager, userRep]
  );

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
          const seedRows = SEED_CLIENTS.map(c => clientToDb({ ...c, id: undefined }));
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

        // Load prospects
        const { data: prospectRows, error: prospectErr } = await supabase
          .from('prospects').select('*').order('created_at');
        if (!prospectErr && prospectRows) {
          setProspects(prospectRows.map(prospectFromDb));
        }

        // Load B2B Customs production tracker rows
        const { data: customsRows, error: customsErr } = await supabase
          .from('b2b_customs').select('*').order('created_at');
        if (!customsErr && customsRows) {
          setB2bCustoms(customsRows.map(b2bCustomFromDb));
        } else if (customsErr) {
          console.warn('[b2b_customs] table missing or not yet created:', customsErr.message);
        }
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

    // Try with all columns first; fall back to base columns if schema is missing extended ones
    let inserted, error;
    ({ data: inserted, error } = await supabase.from('visits').insert(visitToDb(enriched, true)).select().single());
    if (error && error.message && error.message.includes('column')) {
      console.warn('[addVisit] extended columns missing, retrying with base columns only');
      ({ data: inserted, error } = await supabase.from('visits').insert(visitToDb(enriched, false)).select().single());
    }
    if (error) { console.error('[addVisit]', error); throw new Error(error.message); }

    const newVisit = visitFromDb(inserted);
    // Merge in the fields that weren't saved to DB yet so UI shows them correctly
    const fullVisit = { ...newVisit, channel: visit.channel || '', contactMethod: visit.contactMethod || '', followUpDate: visit.followUpDate || '', taggedReps: visit.taggedReps || [] };
    setVisits((prev) => [...prev, fullVisit]);

    // Update client total_sales in DB and local state
    const client = clients.find(c => Number(c.id) === Number(visit.clientId));
    if (client) {
      const newSales = (client.totalSales || 0) + computedTotal;
      // Auto-convert rules:
      // 1. Sold In outcome → always Converted
      // 2. Has SKU items (real sale) → Converted
      // 3. Rejected → Lost
      // 4. New client → Contacted
      // 5. Otherwise keep existing status
      const hasSale = computedTotal > 0 || (visit.items && visit.items.some(it => Number(it.qty) > 0));
      const newStatus = visit.outcome === 'Sold In' ? 'Converted'
        : hasSale ? 'Converted'
        : visit.outcome === 'Rejected' ? 'Lost'
        : (client.status === 'New' ? 'Contacted' : client.status);
      await supabase.from('clients').update({
        total_sales: newSales,
        last_contacted: visit.date,
        status: newStatus,
      }).eq('id', Number(visit.clientId));
      setClients((prev) => prev.map(c => Number(c.id) === Number(visit.clientId)
        ? { ...c, totalSales: newSales, lastContacted: visit.date, status: newStatus }
        : c));
    }
    return fullVisit;
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
    const idNum = Number(visitId);
    const isFakeId = !visitId || isNaN(idNum) || String(visitId).startsWith('seed-');

    // For fake/seed IDs — just remove from local state, nothing in Supabase to delete
    if (isFakeId) {
      setVisits((prev) => prev.filter(x => String(x.id) !== String(visitId)));
      return;
    }

    const v = visits.find(x => Number(x.id) === idNum);
    if (!v) {
      // Not in local state either — nothing to do
      console.warn('[deleteVisit] not found locally:', visitId);
      return;
    }
    const refund = Number(v.saleAmount) || 0;
    const clientId = Number(v.clientId);

    const { error } = await supabase.from('visits').delete().eq('id', idNum);
    if (error) {
      // If it's a "not found" error — record never made it to DB, just clean up local state
      console.warn('[deleteVisit] Supabase error (removing locally anyway):', error.message);
    }

    // Always clean up local state
    setVisits((prev) => prev.filter(x => Number(x.id) !== idNum));

    // Update client totalSales only if there was real revenue
    const client = clients.find(c => Number(c.id) === clientId);
    if (client && refund > 0) {
      const newTotal = Math.max(0, (client.totalSales || 0) - refund);
      await supabase.from('clients').update({ total_sales: newTotal }).eq('id', clientId);
      setClients((prev) => prev.map(c =>
        Number(c.id) === clientId ? { ...c, totalSales: newTotal } : c
      ));
    }
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
    if (patch.paymentTerms !== undefined) dbPatch.payment_terms = patch.paymentTerms;
    if (patch.clientTags !== undefined) dbPatch.client_tags = patch.clientTags;
    if (patch.prospectedAmount !== undefined) dbPatch.prospected_amount = patch.prospectedAmount;
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
    const dbData = clientToDb({ ...newClientData, id: undefined });
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

  // ------- Prospect operations -------
  const addProspect = async (p) => {
    const { data: inserted, error } = await supabase
      .from('prospects').insert(prospectToDb(p)).select().single();
    if (error) { console.error('[addProspect]', error); throw new Error(error.message); }
    setProspects(prev => [...prev, prospectFromDb(inserted)]);
    return prospectFromDb(inserted);
  };

  const updateProspect = async (id, patch) => {
    const dbPatch = {};
    if (patch.clientName !== undefined) dbPatch.client_name = patch.clientName;
    if (patch.clientId !== undefined) dbPatch.client_id = patch.clientId ? Number(patch.clientId) : null;
    if (patch.skuNotes !== undefined) dbPatch.sku_notes = patch.skuNotes;
    if (patch.amount !== undefined) dbPatch.amount = Number(patch.amount) || 0;
    if (patch.expectedDate !== undefined) dbPatch.expected_date = patch.expectedDate || null;
    if (patch.status !== undefined) dbPatch.status = patch.status;
    if (patch.salesRep !== undefined) dbPatch.sales_rep = patch.salesRep;
    if (patch.items !== undefined) dbPatch.items = patch.items;
    const { error } = await supabase.from('prospects').update(dbPatch).eq('id', id);
    if (error) { console.error('[updateProspect]', error); return; }
    setProspects(prev => prev.map(p => p.id === id ? { ...p, ...patch } : p));
  };

  const deleteProspect = async (id) => {
    await supabase.from('prospects').delete().eq('id', id);
    setProspects(prev => prev.filter(p => p.id !== id));
  };

  // ------- B2B Customs (production tracker) operations -------
  const addB2bCustom = async (row) => {
    const { data: inserted, error } = await supabase
      .from('b2b_customs').insert(b2bCustomToDb(row)).select().single();
    if (error) { console.error('[addB2bCustom]', error); throw new Error(error.message); }
    const newRow = b2bCustomFromDb(inserted);
    setB2bCustoms(prev => [...prev, newRow]);
    return newRow;
  };

  const updateB2bCustom = async (id, patch) => {
    const dbPatch = {};
    if (patch.customerName !== undefined)    dbPatch.customer_name    = patch.customerName;
    if (patch.depositPaid !== undefined)     dbPatch.deposit_paid     = patch.depositPaid;
    if (patch.briefed !== undefined)         dbPatch.briefed          = patch.briefed;
    if (patch.liquidLinedUp !== undefined)   dbPatch.liquid_lined_up  = patch.liquidLinedUp;
    if (patch.productionStage !== undefined) dbPatch.production_stage = patch.productionStage;
    if (patch.balancePaid !== undefined)     dbPatch.balance_paid     = patch.balancePaid;
    if (patch.readyDispatch !== undefined)   dbPatch.ready_dispatch   = patch.readyDispatch;
    // Optimistic local update first for snappy typing, then persist
    setB2bCustoms(prev => prev.map(r => r.id === id ? { ...r, ...patch } : r));
    const { error } = await supabase.from('b2b_customs').update(dbPatch).eq('id', id);
    if (error) console.error('[updateB2bCustom]', error);
  };

  const deleteB2bCustom = async (id) => {
    setB2bCustoms(prev => prev.filter(r => r.id !== id));
    const { error } = await supabase.from('b2b_customs').delete().eq('id', id);
    if (error) console.error('[deleteB2bCustom]', error);
  };

  const deleteClient = async (id) => {
    // Delete related visits first to avoid FK constraint violations
    await supabase.from('visits').delete().eq('client_id', id);
    const { error } = await supabase.from('clients').delete().eq('id', id);
    if (error) { console.error('[deleteClient] error:', error); throw new Error(error.message); }
    setClients(prev => prev.filter(c => c.id !== id));
    setVisits(prev => prev.filter(v => v.clientId !== id));
  };

  const monthVisits = useMemo(() => {
    return visits.filter(v => v.date && v.date.startsWith(activeMonth));
  }, [visits, activeMonth]);

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: '#002855' }}>
        <div className="text-center">
          <div className="inline-block animate-pulse">
            <AvanteLogo height={140} />
          </div>
          <p className="mt-6 text-sm tracking-[0.4em]" style={{ color: '#DBB85E', fontFamily: "'Cinzel', serif", fontWeight: 600 }}>LOADING CRM</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen overflow-x-hidden" style={{ background: '#FCF7F2', fontFamily: "'Libre Baskerville', Georgia, serif", color: '#002855' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Libre+Baskerville:ital,wght@0,400;0,700;1,400&family=Cinzel:wght@500;600;700;900&display=swap');
        @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
        *, *::before, *::after { box-sizing: border-box; }
        html, body { margin: 0; padding: 0; overflow-x: hidden; width: 100%; background: #FCF7F2; }
        .font-display { font-family: 'Cinzel', 'Copperplate', serif; letter-spacing: 0.08em; }
        .font-body { font-family: 'Libre Baskerville', Georgia, serif; }
        /* Colour utilities */
        .ink { color: #002855; }
        .ocean { color: #5A7A99; }
        .gold { color: #DBB85E; }
        .copper { color: #BC8D26; }
        .bg-ink { background: #002855; }
        .bg-cream { background: #FCF7F2; }
        .bg-gold { background: #DBB85E; }
        .bg-copper { background: #BC8D26; }
        .border-ink { border-color: #002855; }
        .border-copper { border-color: #BC8D26; }
        .border-gold { border-color: #DBB85E; }
        .diamond-clip { clip-path: polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%); }
        .premium-card { background: #FCF7F2; border: 1px solid rgba(0,40,85,0.15); box-shadow: 0 1px 0 rgba(0,40,85,0.04), 0 8px 24px -16px rgba(0,40,85,0.2); }
        /* Scrollbar */
        .scrollbar-thin::-webkit-scrollbar { width: 4px; height: 4px; }
        .scrollbar-thin::-webkit-scrollbar-track { background: rgba(0,40,85,0.05); }
        .scrollbar-thin::-webkit-scrollbar-thumb { background: rgba(0,40,85,0.3); border-radius: 3px; }
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
          background: #002855;
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
        /* Nav always lives on its own full-width row below the top bar —
           guarantees every tab stays visible and never gets squeezed by the
           logo / user pill / log-visit button competing for the same row. */
        .crm-nav-row {
          width: 100%;
          max-width: 1400px;
          margin: 0 auto;
          border-top: 1px solid rgba(255,255,255,0.08);
        }
        .crm-nav {
          display: flex;
          align-items: center;
          gap: 2px;
          overflow-x: auto;
          -ms-overflow-style: none;
          scrollbar-width: none;
          -webkit-overflow-scrolling: touch;
          padding: 0 8px;
        }
        .crm-nav::-webkit-scrollbar { display: none; }
        .crm-nav-btn {
          display: flex;
          align-items: center;
          gap: 5px;
          padding: 10px 12px;
          font-family: 'Cinzel', 'Copperplate', serif;
          font-size: 10px;
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
        .crm-nav-btn.active { background: #DBB85E; color: #002855; }
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
          background: #BC8D26;
          color: #FCF7F2;
          border: none;
          cursor: pointer;
          white-space: nowrap;
        }
        .crm-log-btn:hover { background: #DBB85E; color: #002855; }
        @media (min-width: 900px) {
          .crm-header-inner { padding: 0 24px; height: 60px; }
          .crm-brand-text { display: block; }
          .crm-nav { padding: 0 16px; gap: 4px; justify-content: center; }
          .crm-nav-btn { font-size: 10px; padding: 9px 16px; }
          .crm-log-btn { font-size: 10px; padding: 9px 18px; }
        }
        @media (min-width: 1024px) {
          .crm-header-inner { padding: 0 40px; }
          .crm-nav-row { padding: 0 24px; }
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
          .grid-kpi { display: grid; grid-template-columns: repeat(3,1fr); gap: 12px; }
          .grid-6 { grid-template-columns: repeat(6,1fr); }
          .grid-lb { display: grid; grid-template-columns: 2fr 1fr; gap: 16px; }
          .grid-repcard { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; }
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
        .border { border: 1px solid rgba(0,40,85,0.2); }
        .border-2 { border: 2px solid; }
        .border-b { border-bottom: 1px solid rgba(0,40,85,0.15); }
        .border-t { border-top: 1px solid rgba(0,40,85,0.15); }
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
        input:focus, select:focus, textarea:focus { outline: none; border-color: #BC8D26 !important; }
        /* Mobile form inputs — prevent zoom on iOS */
        @media (max-width: 639px) {
          input, select, textarea { font-size: 16px !important; }
        }
        /* Hover states */
        .hover-bg:hover { background: rgba(0,40,85,0.06); }
      `}</style>

      <Header view={view} setView={setView} onLog={() => setShowLogModal(true)}
        visits={visits.map(v => ({ ...v, clientName: clients.find(c => c.id === v.clientId)?.venue || '' }))}
        clients={clients}
        currentUser={currentUser}
        onLogout={onLogout}
        onNavigate={(targetView, clientId) => {
          setView(targetView);
          if (clientId) {
            const c = clients.find(cl => cl.id === clientId);
            if (c) setSelectedClient(c);
          }
        }} />

      <main className="crm-main">
        {view === 'notifications' && (
          <NotificationsPage
            visits={visits.map(v => ({ ...v, clientName: clients.find(c => c.id === v.clientId)?.venue || '' }))}
            clients={userIsManager ? clients : visibleClients}
            onSelectClient={(clientId) => {
              const c = clients.find(cl => cl.id === clientId);
              if (c) setSelectedClient(c);
            }}
          />
        )}

        {view === 'dashboard' && (
          <Dashboard
            clients={clients}
            visits={monthVisits}
            allVisits={visits}
            targets={targets}
            activeRep={activeRep}
            setActiveRep={setActiveRep}
            activeMonth={activeMonth}
            setActiveMonth={setActiveMonth}
            prospects={prospects}
            onAddProspect={addProspect}
            onUpdateProspect={updateProspect}
            onDeleteProspect={deleteProspect}
            skuOverrides={skuPrices}
            allowedReps={userIsManager ? ['All', ...SALES_REPS] : ['All', userRep]}
            onNavigate={(v, clientId) => {
              setView(v);
              if (clientId) {
                const c = clients.find(cl => cl.id === clientId);
                if (c) setSelectedClient(c);
              }
            }}
          />
        )}
        {view === 'leads' && (
          <LeadsPage
            clients={visibleClients}
            visits={visits}
            updateClient={updateClient}
            onSelect={setSelectedClient}
            onAddNew={() => setNewClientCtx({ defaultRep: userIsManager ? (activeRep === 'All' ? '' : activeRep) : userRep, onCreated: null })}
            onDelete={async (id) => {
              try {
                const c = clients.find(cl => cl.id === id);
                await deleteClient(id);
                showToast(`Client removed: ${c?.venue || id}`);
              } catch (err) {
                console.error('[LeadsPage] delete failed:', err);
                showToast('Could not delete client — please try again.');
              }
            }}
            onNavigate={(v, clientId) => {
              setView(v);
              if (clientId) { const c = clients.find(cl => cl.id === clientId); if (c) setSelectedClient(c); }
            }}
            defaultRepFilter={userIsManager ? 'All' : userRep}
            lockRepFilter={!userIsManager}
          />
        )}
        {view === 'orders' && (
          <OrderHistoryPage
            clients={clients}
            visits={visits}
            onDeleteVisit={(visitId) => {
              const v = visits.find(x => x.id === visitId);
              const c = v ? clients.find(cl => cl.id === v.clientId) : null;
              const saleNote = v?.saleAmount > 0 ? ` Sale of ${ZAR(Math.round(v.saleAmount))} will be removed.` : '';
              askConfirm({
                title: 'Delete this order?',
                message: `${c?.venue || 'Unknown'} · ${v?.date || ''}${saleNote}\n\nThis cannot be undone.`,
                confirmLabel: 'DELETE ORDER',
                onConfirm: async () => {
                  try {
                    await deleteVisit(visitId);
                    showToast('Order deleted');
                  } catch (err) {
                    console.error('[deleteOrder]', err);
                    showToast('Could not delete order — please try again.');
                  }
                },
              });
            }}
          />
        )}
        {view === 'visits' && (
          <VisitsPage
            visits={visits}
            clients={clients}
            onLog={() => setShowLogModal(true)}
            onEdit={(visit) => setEditingVisit(visit)}
            onEmail={(visit) => setEditingVisit({ ...visit, _emailOnly: true })}
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
                onConfirm: async () => {
                  try {
                    await deleteVisit(visitId);
                    showToast('Visit deleted · client totals updated', 'success');
                  } catch (err) {
                    console.error('[deleteVisit]', err);
                    showToast('Could not delete visit — please try again.');
                  }
                },
              });
            }}
          />
        )}
        {view === 'b2bcustoms' && (
          <B2BCustomsPage
            rows={b2bCustoms}
            onAdd={addB2bCustom}
            onUpdate={updateB2bCustom}
            onDelete={deleteB2bCustom}
            askConfirm={askConfirm}
          />
        )}
        {view === 'manager' && userIsManager && (
          <ManagerPortal
            targets={targets}
            saveTargets={persistTargets}
            clients={clients}
            visits={visits}
            askConfirm={askConfirm}
            skuPrices={skuPrices}
            saveSkuPrices={(prices) => { setSkuPrices(prices); showToast('SKU prices updated'); }}
          />
        )}
      </main>

      {showLogModal && (
        <LogVisitModal
          clients={visibleClients}
          skuOverrides={skuPrices}
          preselectedSalesRep={userIsManager ? '' : userRep}
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
          onDelete={async () => {
            await deleteClient(selectedClient.id);
            setSelectedClient(null);
            showToast(`Client removed: ${selectedClient.venue}`);
          }}
          onDeleteVisit={(visitId) => {
            const v = visits.find(x => x.id === visitId);
            const saleNote = v?.saleAmount > 0 ? `\nSale of ${ZAR(Math.round(v.saleAmount))} will be removed.` : '';
            askConfirm({
              title: 'Delete this visit?',
              message: `${selectedClient.venue} · ${v?.date || ''}${saleNote}\n\nThis cannot be undone.`,
              confirmLabel: 'DELETE VISIT',
              onConfirm: async () => { await deleteVisit(visitId); showToast('Visit deleted'); },
            });
          }}
          onPlaceOrder={(c) => {
            // Keep the client detail popup open — just open the order modal on top
            setPlaceOrderClient(c);
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
          <div className="flex items-center gap-3 px-4 py-3 shadow-2xl border-l-4" style={{ background: '#002855', borderLeftColor: toast.type === 'error' ? '#CC233A' : '#DBB85E', minWidth: '240px', maxWidth: '320px' }}>
            <div className="w-2 h-2 flex-shrink-0 diamond-clip" style={{ background: toast.type === 'error' ? '#CC233A' : '#DBB85E' }}></div>
            <div className="min-w-0">
              <p className="font-display text-[9px] tracking-[0.3em]" style={{ color: '#DBB85E', fontWeight: 600 }}>{toast.type === 'error' ? 'ERROR' : 'SAVED'}</p>
              <p className="text-xs truncate" style={{ color: '#FCF7F2', fontWeight: 500 }}>{toast.message}</p>
            </div>
          </div>
        </div>
      )}

      {/* Place Order modal — sits on top of the client detail popup */}
      {placeOrderClient && (
        <LogVisitModal
          clients={clients}
          skuOverrides={skuPrices}
          preselectedClientId={placeOrderClient.id}
          preselectedSalesRep={placeOrderClient.accountManager || ''}
          onClose={() => setPlaceOrderClient(null)}
          onSubmit={async (v) => {
            await addVisit(v);
            setPlaceOrderClient(null);
            showToast(`Order logged for ${placeOrderClient.venue}`);
          }}
          onRequestNewClient={(rep, cb) => {
            setPlaceOrderClient(null);
            setNewClientCtx({ defaultRep: rep, onCreated: cb });
          }}
        />
      )}

      {confirmCtx && (
        <ConfirmModal
          title={confirmCtx.title}
          message={confirmCtx.message}
          confirmLabel={confirmCtx.confirmLabel}
          danger={confirmCtx.danger}
          onCancel={() => setConfirmCtx(null)}
          onConfirm={async () => {
            const fn = confirmCtx.onConfirm;
            setConfirmCtx(null);
            if (fn) await fn();
          }}
        />
      )}
    </div>
  );
}

// =================== Confirm Modal ===================
function ConfirmModal({ title, message, confirmLabel, danger, onCancel, onConfirm }) {
  return (
    <div style={{ position:"fixed", top:0, left:0, right:0, bottom:0, zIndex:9999, display:"flex", alignItems:"center", justifyContent:"center", padding:20, background:"rgba(0,40,85,0.82)", backdropFilter:"blur(4px)" }} onClick={onCancel}>
      <div style={{ background:'#FCF7F2', maxWidth:360, width:'100%', border:'2px solid #BC8D26', boxShadow:'0 20px 60px rgba(0,0,0,0.4)' }} onClick={(e) => e.stopPropagation()}>
        <div style={{ padding:'12px 16px', background:'#002855', display:'flex', alignItems:'center', gap:10 }}>
          <div style={{ width:8, height:8, background: danger ? '#CC233A' : '#DBB85E', flexShrink:0, transform:'rotate(45deg)' }}></div>
          <div>
            <p style={{ fontFamily:"'Cinzel',serif", fontSize:8, letterSpacing:'0.4em', color:'#DBB85E', fontWeight:600 }}>{danger ? 'DESTRUCTIVE ACTION' : 'CONFIRM'}</p>
            <h2 style={{ fontFamily:"'Cinzel',serif", color:'#FCF7F2', fontWeight:700, fontSize:14, letterSpacing:'0.06em', marginTop:2 }}>{title}</h2>
          </div>
        </div>
        <div style={{ padding:'14px 16px' }}>
          <p style={{ fontSize:12, color:'#002855', lineHeight:1.5, whiteSpace:'pre-line' }}>{message}</p>
        </div>
        <div style={{ display:'flex', justifyContent:'flex-end', gap:8, padding:'0 16px 14px' }}>
          <button type="button" onClick={onCancel}
            style={{ padding:'7px 16px', border:'1px solid rgba(0,40,85,0.2)', background:'none', fontFamily:"'Cinzel',serif", fontSize:9, letterSpacing:'0.2em', color:'#002855', cursor:'pointer', fontWeight:600 }}>
            CANCEL
          </button>
          <button type="button" onClick={onConfirm} autoFocus
            style={{ padding:'7px 16px', background: danger ? '#CC233A' : '#002855', border:'none', fontFamily:"'Cinzel',serif", fontSize:9, letterSpacing:'0.2em', color:'#FCF7F2', cursor:'pointer', fontWeight:700 }}>
            {confirmLabel || 'CONFIRM'}
          </button>
        </div>
      </div>
    </div>
  );
}

// =================== Header ===================
function Header({ view, setView, onLog, visits, clients, currentUser, onLogout, onNavigate }) {
  const userIsManager = isManager(currentUser);

  const tabs = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'leads', label: 'Leads & Clients', icon: Users },
    { id: 'orders', label: 'Client Order History', icon: ShoppingCart },
    { id: 'visits', label: 'Visit Log', icon: ClipboardList },
    { id: 'b2bcustoms', label: 'B2B Customs', icon: Briefcase },
    // Manager tab only visible to manager
    ...(userIsManager ? [{ id: 'manager', label: 'Manager', icon: Settings }] : []),
  ];

  // Count notifications
  const notifCount = useMemo(() => {
    let count = 0;
    const today = new Date().toISOString().slice(0, 10);
    // Tagged visit notifications
    (visits || []).forEach(v => {
      if (v.taggedReps && v.taggedReps.length > 0) {
        if (userIsManager || v.taggedReps.includes(currentUser?.rep)) count++;
      }
      // Follow-up due today or overdue — notify the client's account manager
      if (v.followUpDate && v.followUpDate <= today) {
        const client = (clients || []).find(c => Number(c.id) === Number(v.clientId));
        if (client) {
          if (userIsManager || client.accountManager === currentUser?.rep) count++;
        }
      }
    });
    (clients || []).forEach(c => {
      if (c.clientTags && c.clientTags.length > 0) {
        if (userIsManager || c.clientTags.includes(currentUser?.rep)) count++;
      }
    });
    return count;
  }, [visits, clients, userIsManager, currentUser]);
  return (
    <header className="crm-header">
      <div className="crm-header-inner">
        {/* Logo */}
        <div className="crm-logo-wrap">
          <AvanteLogo height={40} />
          <div className="crm-brand-text" style={{ borderLeft: '1px solid rgba(255,255,255,0.2)', paddingLeft: 12 }}>
            <p className="font-display" style={{ fontSize: 8, letterSpacing: '0.35em', color: '#DBB85E', fontWeight: 600, margin: 0 }}>SALES CRM</p>
            <p className="font-display" style={{ fontSize: 7, letterSpacing: '0.2em', color: '#FCF7F2', opacity: 0.5, margin: '2px 0 0' }}>FORWARD TOGETHER</p>
          </div>
        </div>
        {/* Right side — Notifications + user + Log Visit */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0, marginLeft: 'auto' }}>
          <button
            onClick={() => setView('notifications')}
            style={{ position: 'relative', background: view === 'notifications' ? 'rgba(219,184,94,0.2)' : notifCount > 0 ? 'rgba(219,184,94,0.1)' : 'none', border: notifCount > 0 ? '1px solid #DBB85E' : '1px solid rgba(255,255,255,0.2)', borderRadius: 4, padding: '6px 10px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, color: notifCount > 0 ? '#DBB85E' : 'rgba(255,255,255,0.5)', transition: 'all 0.2s' }}
            title="Notifications">
            <Bell style={{ width: 15, height: 15 }} />
            {notifCount > 0 && (
              <span style={{ background: '#CC233A', color: '#FCF7F2', fontFamily: "'Cinzel',serif", fontSize: 9, fontWeight: 700, borderRadius: 10, padding: '1px 6px', minWidth: 16, textAlign: 'center', lineHeight: 1.6 }}>
                {notifCount > 99 ? '99+' : notifCount}
              </span>
            )}
          </button>
          {/* User name pill */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '5px 10px', border: '1px solid rgba(255,255,255,0.15)', borderRadius: 4, background: 'rgba(255,255,255,0.06)' }}>
            <span style={{ fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.15em', fontWeight: 700, color: '#DBB85E' }}>{currentUser?.name?.toUpperCase()}</span>
            {userIsManager && <span style={{ fontSize: 8, color: 'rgba(255,255,255,0.4)', fontFamily: "'Cinzel',serif", letterSpacing: '0.1em' }}>MANAGER</span>}
          </div>
          {/* Logout */}
          <button onClick={onLogout} title="Sign out"
            style={{ background: 'none', border: '1px solid rgba(255,255,255,0.15)', borderRadius: 4, padding: '6px 8px', cursor: 'pointer', color: 'rgba(255,255,255,0.5)', display: 'flex', alignItems: 'center', transition: 'all 0.2s' }}
            onMouseEnter={e => { e.currentTarget.style.color = '#CC233A'; e.currentTarget.style.borderColor = '#CC233A'; }}
            onMouseLeave={e => { e.currentTarget.style.color = 'rgba(255,255,255,0.5)'; e.currentTarget.style.borderColor = 'rgba(255,255,255,0.15)'; }}>
            <LogOut style={{ width: 14, height: 14 }} />
          </button>
          <button onClick={onLog} className="crm-log-btn">
            <Plus style={{ width: 14, height: 14 }} />
            <span>LOG VISIT</span>
          </button>
        </div>
      </div>

      {/* Nav row — full width, own row, always visible at every screen size so tabs never get squeezed */}
      <div className="crm-nav-row">
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
      </div>
    </header>
  );
}

// =================== Dashboard ===================
// ── OVERDUE CLIENTS KANBAN ──────────────────────────────────────────────────
function OverdueClients({ clients, visits, onNavigate, visibleReps }) {
  const today = new Date();
  const repsToShow = visibleReps || SALES_REPS;

  // Build top-5 overdue list per rep
  const overdueByRep = useMemo(() => {
    const result = {};
    repsToShow.forEach(rep => {
      result[rep] = clients
        .filter(c => c.accountManager === rep && c.status !== 'Lost')
        .map(c => {
          const clientVisits = visits.filter(v => v.clientId === c.id);
          const latestVisit = clientVisits.sort((a, b) => (b.date || '').localeCompare(a.date || ''))[0];
          const lastDate = latestVisit?.date || c.lastContacted || null;
          const daysAgo = lastDate
            ? Math.floor((today.getTime() - new Date(lastDate).getTime()) / (1000 * 60 * 60 * 24))
            : 999;
          return { ...c, lastDate, daysAgo };
        })
        .filter(c => c.daysAgo >= 30)
        .sort((a, b) => b.daysAgo - a.daysAgo)
        .slice(0, 15);
    });
    return result;
  }, [clients, visits]);

  const totalOverdue = SALES_REPS.reduce((s, r) => s + (overdueByRep[r]?.length || 0), 0);

  return (
    <div className="premium-card p-4">
      {/* Header */}
      <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
        <div className="flex items-center gap-2">
          <div className="w-2 h-2 diamond-clip" style={{ background: '#CC233A' }}></div>
          <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>
            OVERDUE FOLLOW-UPS
          </h2>
        </div>
        <span className="font-display text-[9px] tracking-[0.15em]" style={{ color: '#CC233A', fontWeight: 600 }}>
          {totalOverdue} CLIENT{totalOverdue !== 1 ? 'S' : ''} · &gt;30 DAYS
        </span>
      </div>

      {/* Kanban — 3 reps per row, cards shown 3-across within each rep column */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 12 }}>
        {repsToShow.map(rep => {
          const cards = overdueByRep[rep] || [];
          return (
            <div key={rep} style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
              {/* Column header */}
              <div style={{ padding: '8px 10px', background: '#002855', display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 6 }}>
                <span style={{ fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.2em', fontWeight: 700, color: '#FCF7F2' }}>{rep.toUpperCase()}</span>
                {cards.length > 0
                  ? <span style={{ background: '#CC233A', color: '#FCF7F2', fontFamily: "'Cinzel',serif", fontSize: 8, fontWeight: 700, borderRadius: 10, padding: '1px 7px' }}>{cards.length}</span>
                  : <span style={{ background: '#2d8659', color: '#FCF7F2', fontFamily: "'Cinzel',serif", fontSize: 8, fontWeight: 700, borderRadius: 10, padding: '1px 7px' }}>✓</span>
                }
              </div>

              {/* Cards — 3 across × 5 rows */}
              {cards.length === 0 ? (
                <div style={{ padding: '14px 8px', textAlign: 'center', background: 'rgba(45,134,89,0.06)', border: '1px solid rgba(45,134,89,0.15)', borderRadius: 2 }}>
                  <p style={{ fontSize: 10, color: '#2d8659', fontStyle: 'italic' }}>Up to date 🎉</p>
                </div>
              ) : (
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 5 }}>
                  {cards.map((c, i) => (
                    <div key={c.id}
                      onClick={() => onNavigate && onNavigate(c.id)}
                      style={{ background: '#FCF7F2', border: '1px solid rgba(0,40,85,0.1)', borderTop: `3px solid ${c.daysAgo >= 60 ? '#CC233A' : '#BC8D26'}`, padding: '7px 8px', cursor: 'pointer', transition: 'box-shadow 0.15s' }}
                      onMouseEnter={e => e.currentTarget.style.boxShadow = '0 2px 8px rgba(0,40,85,0.12)'}
                      onMouseLeave={e => e.currentTarget.style.boxShadow = 'none'}>
                      <p className="font-display ink" style={{ fontWeight: 700, fontSize: 10, lineHeight: 1.3, overflow: 'hidden', textOverflow: 'ellipsis', display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', marginBottom: 4 }}>{c.venue}</p>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <span style={{ padding: '1px 4px', background: getStatusColor(c.status), color: '#FCF7F2', fontFamily: "'Cinzel',serif", fontSize: 7, letterSpacing: '0.05em', fontWeight: 700 }}>
                          {(c.status || 'NEW').substring(0, 4).toUpperCase()}
                        </span>
                        <span style={{ fontSize: 10, fontWeight: 700, color: c.daysAgo >= 60 ? '#CC233A' : '#BC8D26', fontFamily: "'Cinzel',serif" }}>
                          {c.daysAgo === 999 ? '∞' : `${c.daysAgo}d`}
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// =================== Prospect / Pipeline Forecast Widget ===================
function ProspectWidget({ activeRep = 'All', targets = {}, clients = [], onNavigate }) {
  // B2B clients only, filtered by rep
  const b2bClients = useMemo(() => {
    return clients.filter(c => {
      if (c.channel !== 'B2B') return false;
      if (activeRep !== 'All' && c.accountManager !== activeRep) return false;
      return true;
    }).sort((a, b) => (b.prospectedAmount || 0) - (a.prospectedAmount || 0));
  }, [clients, activeRep]);

  const totalPipeline = useMemo(
    () => b2bClients.reduce((s, c) => s + (Number(c.prospectedAmount) || 0), 0),
    [b2bClients]
  );

  const withAmount = b2bClients.filter(c => Number(c.prospectedAmount) > 0);

  // Build 6-month forecast
  // Rule: if status === 'Invoiced / Paid' → 50% this month, 50% next month
  //        otherwise → full amount sits in current month (pipeline view)
  const sixMonthForecast = useMemo(() => {
    const now = new Date();
    // Build 6 month buckets starting from current month
    const months = Array.from({ length: 6 }, (_, i) => {
      const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
      return {
        label: d.toLocaleString('default', { month: 'short' }).toUpperCase(),
        year: d.getFullYear(),
        monthIdx: d.getMonth(),
        amount: 0,
      };
    });

    b2bClients.forEach(c => {
      const amt = Number(c.prospectedAmount) || 0;
      if (!amt) return;

      if (c.status === 'Invoiced / Paid') {
        // 50% this month, 50% next month
        if (months[0]) months[0].amount += amt * 0.5;
        if (months[1]) months[1].amount += amt * 0.5;
      }
      // Other statuses contribute to the pipeline total only (not placed into future months yet)
    });

    return months;
  }, [b2bClients]);

  const maxMonthAmount = Math.max(...sixMonthForecast.map(m => m.amount), 1);

  return (
    <div className="premium-card p-4">
      {/* Header */}
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
        <div className="flex items-center gap-2">
          <div className="w-2 h-2 diamond-clip" style={{ background: '#DBB85E' }}></div>
          <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>
            B2B PIPELINE FORECAST {activeRep !== 'All' && `— ${activeRep.toUpperCase()}`}
          </h2>
        </div>
        <span className="font-display text-[9px] tracking-[0.15em] ocean" style={{ fontWeight: 600 }}>
          {b2bClients.length} B2B client{b2bClients.length !== 1 ? 's' : ''}
        </span>
      </div>

      {/* Total */}
      <div className="flex items-end justify-between gap-4 mb-4">
        <div>
          <p className="font-display text-[10px] tracking-[0.25em] copper" style={{ fontWeight: 600 }}>PROJECTED B2B PIPELINE</p>
          <p className="font-display text-2xl ink mt-1" style={{ fontWeight: 700 }}>{ZAR(totalPipeline)}</p>
          <p className="italic ocean" style={{ fontSize: 10, marginTop: 2 }}>
            {withAmount.length} client{withAmount.length !== 1 ? 's' : ''} with prospected amounts
          </p>
        </div>
      </div>

      {/* 6-Month Revenue Forecast Bar */}
      <div style={{ borderTop: '1px solid rgba(0,40,85,0.1)', paddingTop: 12, marginBottom: 12 }}>
        <p className="font-display text-[9px] tracking-[0.25em] copper mb-3" style={{ fontWeight: 600 }}>6-MONTH REVENUE FORECAST · INVOICED / PAID CLIENTS</p>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(6,1fr)', gap: 6 }}>
          {sixMonthForecast.map((m, i) => {
            const barPct = maxMonthAmount > 0 ? (m.amount / maxMonthAmount) * 100 : 0;
            const isCurrentMonth = i === 0;
            return (
              <div key={i} style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4 }}>
                {/* Bar */}
                <div style={{ width: '100%', height: 60, display: 'flex', alignItems: 'flex-end', background: 'rgba(0,40,85,0.04)', borderRadius: 3, overflow: 'hidden' }}>
                  <div style={{
                    width: '100%',
                    height: `${Math.max(barPct, m.amount > 0 ? 8 : 0)}%`,
                    background: m.amount > 0 ? (isCurrentMonth ? '#BC8D26' : '#DBB85E') : 'transparent',
                    transition: 'height 0.5s',
                    borderRadius: '2px 2px 0 0',
                  }} />
                </div>
                {/* Amount */}
                <p style={{ fontFamily: "'Cinzel',serif", fontSize: 8, fontWeight: 700, color: m.amount > 0 ? '#002855' : 'rgba(0,40,85,0.25)', textAlign: 'center', lineHeight: 1.3 }}>
                  {m.amount > 0 ? ZAR(m.amount) : '—'}
                </p>
                {/* Month label */}
                <p style={{ fontFamily: "'Cinzel',serif", fontSize: 8, fontWeight: 600, color: isCurrentMonth ? '#BC8D26' : '#5A7A99', letterSpacing: '0.1em', textAlign: 'center' }}>
                  {m.label}
                </p>
              </div>
            );
          })}
        </div>
        <p style={{ fontSize: 9, fontStyle: 'italic', color: '#5A7A99', marginTop: 8 }}>
          When a B2B client reaches <strong>Invoiced / Paid</strong> status, 50% of their prospected amount is allocated to the current month and 50% to the following month.
        </p>
      </div>

      {/* Client rows */}
      {b2bClients.length === 0 ? (
        <div style={{ padding: '16px 0', textAlign: 'center', borderTop: '1px solid rgba(0,40,85,0.1)' }}>
          <p className="italic ocean" style={{ fontSize: 12 }}>No B2B clients yet — add a client under the B2B channel to start forecasting.</p>
        </div>
      ) : (
        <div style={{ borderTop: '1px solid rgba(0,40,85,0.1)' }}>
          {b2bClients.slice(0, 8).map((c, i) => {
            const pct = getB2BPct(c.status);
            return (
              <div key={c.id}
                onClick={() => onNavigate && onNavigate('leads', c.id)}
                style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '9px 8px', borderBottom: i < Math.min(b2bClients.length, 8) - 1 ? '1px solid rgba(0,40,85,0.07)' : 'none', cursor: 'pointer', transition: 'background 0.15s' }}
                onMouseEnter={e => e.currentTarget.style.background = 'rgba(0,40,85,0.03)'}
                onMouseLeave={e => e.currentTarget.style.background = 'transparent'}>
                <div style={{ width: 4, alignSelf: 'stretch', background: Number(c.prospectedAmount) > 0 ? '#DBB85E' : 'rgba(0,40,85,0.1)', flexShrink: 0, borderRadius: 2 }}></div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <p className="font-display ink" style={{ fontWeight: 700, fontSize: 12, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{c.venue}</p>
                  <p style={{ fontSize: 10, color: '#5A7A99', fontStyle: 'italic', marginTop: 1 }}>
                    {c.accountManager}{c.status ? ` · ${c.status}` : ''}
                  </p>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0 }}>
                  {pct !== null && (
                    <span style={{ fontFamily: "'Cinzel',serif", fontSize: 10, fontWeight: 700, color: pct >= 80 ? '#2d8659' : pct >= 40 ? '#BC8D26' : '#5A7A99' }}>
                      {pct}%
                    </span>
                  )}
                  <p className="font-display copper" style={{ fontWeight: 700, fontSize: 12 }}>
                    {Number(c.prospectedAmount) > 0 ? ZAR(c.prospectedAmount) : <span style={{ color: 'rgba(0,40,85,0.3)', fontStyle: 'italic', fontSize: 10 }}>No amount</span>}
                  </p>
                </div>
              </div>
            );
          })}
          {b2bClients.length > 8 && (
            <p style={{ textAlign: 'center', fontSize: 10, fontStyle: 'italic', color: '#5A7A99', padding: '8px 0' }}>
              +{b2bClients.length - 8} more B2B clients
            </p>
          )}
        </div>
      )}
    </div>
  );
}

// =================== Notifications Page ===================
function NotificationsPage({ visits, clients, onSelectClient }) {
  const [filter, setFilter] = useState('All'); // All | Visit | Client | FollowUp
  const [repFilter, setRepFilter] = useState('All');
  const today = new Date().toISOString().slice(0, 10);

  const allItems = useMemo(() => {
    const items = [];

    // Visit-level tags
    (visits || []).forEach(v => {
      if (v.taggedReps && v.taggedReps.length > 0) {
        const client = (clients || []).find(c => Number(c.id) === Number(v.clientId));
        items.push({
          key: `visit-${v.id}`,
          type: 'visit',
          clientId: v.clientId,
          clientName: client?.venue || `Client #${v.clientId}`,
          clientChannel: client?.channel || '',
          clientLocation: client?.location || '',
          taggedReps: v.taggedReps,
          date: v.date || '',
          subtitle: `${v.outcome || '—'} · by ${v.salesRep || '—'}`,
          notes: v.notes || '',
          followUp: v.followUp || '',
        });
      }

      // Follow-up due notifications — tagged to the client's account manager
      if (v.followUpDate && v.followUpDate <= today) {
        const client = (clients || []).find(c => Number(c.id) === Number(v.clientId));
        if (client) {
          const daysOverdue = Math.floor((new Date(today) - new Date(v.followUpDate)) / (1000*60*60*24));
          items.push({
            key: `followup-${v.id}`,
            type: 'followup',
            clientId: v.clientId,
            clientName: client?.venue || `Client #${v.clientId}`,
            clientChannel: client?.channel || '',
            clientLocation: client?.location || '',
            taggedReps: [client.accountManager],
            date: v.followUpDate || '',
            subtitle: daysOverdue === 0 ? '⏰ Follow-up due TODAY' : `⚠️ Follow-up overdue by ${daysOverdue} day${daysOverdue !== 1 ? 's' : ''}`,
            notes: v.notes || '',
            followUp: v.followUpDate || '',
            daysOverdue,
          });
        }
      }
    });

    // Client-level tags
    (clients || []).forEach(c => {
      if (c.clientTags && c.clientTags.length > 0) {
        items.push({
          key: `client-${c.id}`,
          type: 'client',
          clientId: c.id,
          clientName: c.venue,
          clientChannel: c.channel || '',
          clientLocation: c.location || '',
          taggedReps: c.clientTags,
          date: c.lastContacted || '',
          subtitle: `${c.channel || '—'} · ${c.accountManager || '—'} · ${c.status || '—'}`,
          notes: c.notes || '',
          followUp: '',
        });
      }
    });

    return items.sort((a, b) => {
      // Follow-ups first, then by date
      if (a.type === 'followup' && b.type !== 'followup') return -1;
      if (b.type === 'followup' && a.type !== 'followup') return 1;
      return (b.date || '').localeCompare(a.date || '');
    });
  }, [visits, clients, today]);

  const filtered = useMemo(() => {
    return allItems.filter(item => {
      if (filter !== 'All' && item.type !== filter.toLowerCase()) return false;
      if (repFilter !== 'All' && !item.taggedReps.includes(repFilter)) return false;
      return true;
    });
  }, [allItems, filter, repFilter]);

  // Group by tagged rep
  const byRep = useMemo(() => {
    if (repFilter !== 'All') return { [repFilter]: filtered };
    const map = {};
    filtered.forEach(item => {
      (item.taggedReps || []).forEach(rep => {
        if (!map[rep]) map[rep] = [];
        if (!map[rep].find(x => x.key === item.key)) map[rep].push(item);
      });
    });
    return map;
  }, [filtered, repFilter]);

  const typeColors = {
    visit:   { bg: 'rgba(45,134,89,0.1)',  color: '#2d8659',  label: 'VISIT' },
    client:  { bg: 'rgba(90,122,153,0.1)', color: '#5A7A99',  label: 'CLIENT' },
    followup:{ bg: 'rgba(204,35,58,0.1)',  color: '#CC233A',  label: '⏰ FOLLOW-UP DUE' },
  };

  return (
    <div className="space-y-4 fade-up">
      {/* Header */}
      <div className="pb-3 border-b">
        <p className="font-display text-[9px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>AGENT MENTIONS</p>
        <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', flexWrap: 'wrap', gap: 12, marginTop: 4 }}>
          <div>
            <h1 className="font-display ink mt-1" style={{ fontWeight: 700, fontSize: 28 }}>NOTIFICATIONS</h1>
            <p className="italic ocean" style={{ fontSize: 12, marginTop: 2 }}>{allItems.length} total tag{allItems.length !== 1 ? 's' : ''} across {Object.keys(byRep).length} agent{Object.keys(byRep).length !== 1 ? 's' : ''}</p>
          </div>
        </div>
      </div>

      {/* Filters */}
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
        {/* Type filter */}
        <div style={{ display: 'flex', border: '1px solid rgba(0,40,85,0.2)', overflow: 'hidden' }}>
          {['All', 'Visit', 'Client', 'Followup'].map(t => (
            <button key={t} onClick={() => setFilter(t)}
              style={{ padding: '7px 14px', fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.15em', fontWeight: 700, border: 'none', background: filter === t ? '#002855' : 'transparent', color: filter === t ? '#FCF7F2' : '#002855', cursor: 'pointer' }}>
              {t.toUpperCase()}
            </button>
          ))}
        </div>
        {/* Rep filter */}
        <div style={{ display: 'flex', border: '1px solid rgba(0,40,85,0.2)', overflow: 'hidden' }}>
          {['All', ...SALES_REPS].map(r => (
            <button key={r} onClick={() => setRepFilter(r)}
              style={{ padding: '7px 14px', fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.15em', fontWeight: 700, border: 'none', background: repFilter === r ? '#BC8D26' : 'transparent', color: repFilter === r ? '#FCF7F2' : '#002855', cursor: 'pointer' }}>
              {r === 'All' ? 'ALL REPS' : `@${r.toUpperCase()}`}
            </button>
          ))}
        </div>
        <p style={{ fontSize: 11, fontStyle: 'italic', color: '#5A7A99', margin: 0 }}>{filtered.length} result{filtered.length !== 1 ? 's' : ''}</p>
      </div>

      {/* Content */}
      {filtered.length === 0 ? (
        <div className="premium-card" style={{ padding: 40, textAlign: 'center' }}>
          <Bell style={{ width: 32, height: 32, color: 'rgba(0,40,85,0.15)', margin: '0 auto 12px', display: 'block' }} />
          <p style={{ fontSize: 14, color: '#5A7A99', fontStyle: 'italic' }}>No tag notifications match your filters.</p>
        </div>
      ) : (
        Object.entries(byRep).map(([rep, items]) => (
          <div key={rep} className="premium-card" style={{ overflow: 'hidden' }}>
            {/* Rep header */}
            <div style={{ padding: '12px 16px', background: '#002855', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <span style={{ padding: '3px 12px', background: '#BC8D26', color: '#FCF7F2', fontFamily: "'Cinzel',serif", fontSize: 10, letterSpacing: '0.2em', fontWeight: 700 }}>@{rep.toUpperCase()}</span>
                <span style={{ color: 'rgba(255,255,255,0.6)', fontSize: 11, fontStyle: 'italic' }}>{items.length} tag{items.length !== 1 ? 's' : ''}</span>
              </div>
            </div>

            {/* Items */}
            {items.map((item, i) => (
              <div key={item.key}
                onClick={() => onSelectClient && onSelectClient(item.clientId)}
                style={{ display: 'flex', alignItems: 'flex-start', gap: 14, padding: '14px 16px', borderBottom: i < items.length - 1 ? '1px solid rgba(0,40,85,0.07)' : 'none', cursor: 'pointer', transition: 'background 0.15s' }}
                onMouseEnter={e => e.currentTarget.style.background = 'rgba(0,40,85,0.03)'}
                onMouseLeave={e => e.currentTarget.style.background = 'transparent'}>

                {/* Type badge */}
                <div style={{ flexShrink: 0, marginTop: 2 }}>
                  <span style={{ padding: '3px 8px', background: typeColors[item.type].bg, color: typeColors[item.type].color, fontFamily: "'Cinzel',serif", fontSize: 8, letterSpacing: '0.12em', fontWeight: 700, border: `1px solid ${typeColors[item.type].color}` }}>
                    {typeColors[item.type].label}
                  </span>
                </div>

                {/* Main content */}
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, flexWrap: 'wrap' }}>
                    <p style={{ fontFamily: "'Cinzel',serif", fontWeight: 700, fontSize: 14, color: '#002855', margin: 0 }}>{item.clientName}</p>
                    {item.date && <span style={{ fontSize: 10, color: '#5A7A99', fontStyle: 'italic' }}>{item.date}</span>}
                    {item.clientLocation && <span style={{ fontSize: 10, color: 'rgba(0,40,85,0.4)' }}>{item.clientLocation}</span>}
                  </div>
                  <p style={{ fontSize: 11, color: '#5A7A99', fontStyle: 'italic', margin: '3px 0 0' }}>{item.subtitle}</p>
                  {item.notes && <p style={{ fontSize: 11, color: '#002855', margin: '6px 0 0', lineHeight: 1.5 }}>📝 {item.notes}</p>}
                  {item.followUp && <p style={{ fontSize: 11, color: '#BC8D26', margin: '4px 0 0' }}>→ {item.followUp}</p>}
                  {/* All tagged reps */}
                  {item.taggedReps.length > 1 && (
                    <div style={{ display: 'flex', gap: 4, marginTop: 6, flexWrap: 'wrap' }}>
                      {item.taggedReps.map(r => (
                        <span key={r} style={{ padding: '2px 7px', background: 'rgba(188,141,38,0.12)', color: '#BC8D26', fontFamily: "'Cinzel',serif", fontSize: 8, letterSpacing: '0.1em', fontWeight: 700 }}>@{r.toUpperCase()}</span>
                      ))}
                    </div>
                  )}
                </div>

                <ChevronRight style={{ width: 16, height: 16, color: '#BC8D26', flexShrink: 0, marginTop: 4 }} />
              </div>
            ))}
          </div>
        ))
      )}
    </div>
  );
}

// ── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({ clients, visits, allVisits, targets, activeRep, setActiveRep, activeMonth, setActiveMonth, prospects, onAddProspect, onUpdateProspect, onDeleteProspect, skuOverrides, onNavigate, allowedReps }) {
  // Parse activeMonth into a Date for display and daily calculations
  const [selYear, selMonthIdx] = activeMonth.split('-').map(Number);
  const selectedDate = new Date(selYear, selMonthIdx - 1, 1);
  const isCurrentMonth = activeMonth === monthISO();

  const repFilter = (item, key = 'accountManager') =>
    activeRep === 'All' ? true : item[key] === activeRep;
  const filteredClients = clients.filter(c => repFilter(c));
  const filteredVisits = visits.filter(v => repFilter(v, 'salesRep'));

  const monthSales = filteredVisits.reduce((s, v) => s + (v.saleAmount || 0), 0);
  const visitCount = filteredVisits.length;
  const sold = filteredVisits.filter(isSoldVisit).length;
  const conversionRate = visitCount > 0 ? Math.round((sold / visitCount) * 100) : 0;

  // Daily data for line graphs — selected month
  const daysInMonth = new Date(selYear, selMonthIdx, 0).getDate();
  const todayDay = isCurrentMonth ? new Date().getDate() : daysInMonth;
  const dailySales = Array.from({ length: daysInMonth }, (_, i) => {
    const day = String(i + 1).padStart(2, '0');
    const dateStr = `${activeMonth}-${day}`;
    return filteredVisits.filter(v => v.date === dateStr).reduce((s, v) => s + (v.saleAmount || 0), 0);
  });
  const dailyVisits = Array.from({ length: daysInMonth }, (_, i) => {
    const day = String(i + 1).padStart(2, '0');
    const dateStr = `${activeMonth}-${day}`;
    return filteredVisits.filter(v => v.date === dateStr).length;
  });
  const dailyConversion = Array.from({ length: daysInMonth }, (_, i) => {
    const day = String(i + 1).padStart(2, '0');
    const dateStr = `${activeMonth}-${day}`;
    const dayVisits = filteredVisits.filter(v => v.date === dateStr);
    const daySold = dayVisits.filter(isSoldVisit).length;
    return dayVisits.length > 0 ? Math.round((daySold / dayVisits.length) * 100) : 0;
  });

  const activeTargets = useMemo(() => {
    if (activeRep === 'All') {
      return SALES_REPS.reduce((acc, r) => ({
        revenue:     acc.revenue     + (targets[r]?.revenue     || 0),
        visits:      acc.visits      + (targets[r]?.visits      || 0),
        tradeRetail: acc.tradeRetail + (targets[r]?.tradeRetail || 0),
        onCon:       acc.onCon       + (targets[r]?.onCon       || 0),
        b2b:         acc.b2b         + (targets[r]?.b2b         || 0),
      }), { revenue: 0, visits: 0, tradeRetail: 0, onCon: 0, b2b: 0 });
    }
    return targets[activeRep] || { revenue: 0, visits: 0, tradeRetail: 0, onCon: 0, b2b: 0 };
  }, [activeRep, targets]);

  const pct = (a, b) => b > 0 ? Math.min(100, Math.round((a / b) * 100)) : 0;

  // Stage counts for pipeline summary
  const stages = ['New', 'Contacted', 'Converted', 'Lost'];
  const stageCounts = Object.fromEntries(stages.map(s => [s, filteredClients.filter(c => c.status === s).length]));
  const stageColors = { New: '#5A7A99', Contacted: '#DBB85E', Converted: '#2d8659', Lost: '#CC233A' };

  return (
    <div className="space-y-4 md:space-y-6 fade-up">

      {/* Page header */}
      <div className="flex items-center justify-between gap-3 flex-wrap pb-3 border-b">
        <div className="flex items-center gap-3">
          {/* Prev month */}
          <button
            onClick={() => {
              const [y, m] = activeMonth.split('-').map(Number);
              const d = new Date(y, m - 2, 1);
              setActiveMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
            }}
            style={{ background: 'none', border: '1px solid rgba(0,40,85,0.2)', cursor: 'pointer', padding: '4px 10px', color: '#002855', fontFamily: "'Cinzel',serif", fontSize: 14, fontWeight: 700 }}
            title="Previous month">‹</button>

          <div>
            <p className="font-display text-[9px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>PERFORMANCE OVERVIEW</p>
            <h1 className="font-display ink mt-0.5" style={{ fontWeight: 700, fontSize: 28 }}>
              {selectedDate.toLocaleDateString('en-ZA', { month: 'long', year: 'numeric' }).toUpperCase()}
            </h1>
            {!isCurrentMonth && (
              <button
                onClick={() => setActiveMonth(monthISO())}
                style={{ marginTop: 4, background: 'none', border: 'none', cursor: 'pointer', fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.2em', color: '#BC8D26', fontWeight: 600, padding: 0 }}>
                ← BACK TO CURRENT MONTH
              </button>
            )}
          </div>

          {/* Next month — only allow up to current month */}
          <button
            onClick={() => {
              const [y, m] = activeMonth.split('-').map(Number);
              const next = new Date(y, m, 1);
              const nextStr = `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, '0')}`;
              if (nextStr <= monthISO()) setActiveMonth(nextStr);
            }}
            style={{ background: 'none', border: '1px solid rgba(0,40,85,0.2)', cursor: isCurrentMonth ? 'not-allowed' : 'pointer', padding: '4px 10px', color: isCurrentMonth ? 'rgba(0,40,85,0.3)' : '#002855', fontFamily: "'Cinzel',serif", fontSize: 14, fontWeight: 700 }}
            title="Next month"
            disabled={isCurrentMonth}>›</button>
        </div>
        <RepToggle active={activeRep} onChange={setActiveRep} allowedReps={allowedReps} />
      </div>

      {/* ── KPI CARDS ── */}
      <div className="grid-kpi">
        <KPICard label="Month Sales" value={ZAR(monthSales)} target={activeTargets.revenue} targetValue={ZAR(activeTargets.revenue)} progress={pct(monthSales, activeTargets.revenue)} icon={DollarSign} accent="copper" sparkData={dailySales} todayDay={todayDay} />
        <KPICard label="Visits" value={visitCount} target={activeTargets.visits} targetValue={`${activeTargets.visits} visits`} progress={pct(visitCount, activeTargets.visits)} icon={Activity} accent="ocean" sparkData={dailyVisits} todayDay={todayDay} />
        <KPICard label="Conversion" value={`${conversionRate}%`} subtitle={`${sold} of ${visitCount} sold`} progress={conversionRate} icon={TrendingUp} accent="gold" sparkData={dailyConversion} todayDay={todayDay} />
      </div>

      {/* ── PIPELINE FORECAST ── */}
      <ProspectWidget
        activeRep={activeRep}
        targets={targets}
        clients={clients}
        onNavigate={onNavigate}
      />

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
          { id: 'leads', label: 'Leads & Clients', sub: `${filteredClients.length} venues`, icon: Users, color: '#5A7A99' },
          { id: 'visits', label: 'Visit Log', sub: `${filteredVisits.length} this month`, icon: ClipboardList, color: '#BC8D26' },
          { id: 'manager', label: 'Manager Portal', sub: 'Targets & Export', icon: Settings, color: '#002855' },
        ].map(item => {
          const Icon = item.icon;
          return (
            <button key={item.id} onClick={() => onNavigate(item.id)}
              className="premium-card p-4 text-left hover:shadow-md transition-all active:scale-[0.98] flex items-center gap-3">
              <div className="w-10 h-10 flex-shrink-0 flex items-center justify-center" style={{ background: item.color }}>
                <Icon className="w-5 h-5" style={{ color: '#FCF7F2' }} />
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

function RepToggle({ active, onChange, allowedReps }) {
  const opts = allowedReps || ['All', ...SALES_REPS];
  return (
    <div className="inline-flex border border bg-cream">
      {opts.map((r, i) => {
        const isActive = active === r;
        return (
          <button
            key={r}
            onClick={() => onChange(r)}
            className={`px-2.5 md:px-4 py-1.5 md:py-2 font-display text-[9px] md:text-[11px] tracking-[0.2em] transition-all ${isActive ? 'bg-ink' : 'hover:'} ${i > 0 ? 'border-l border' : ''}`}
            style={{ fontWeight: 600, color: isActive ? '#FCF7F2' : '#002855' }}
          >
            {r === 'All' ? 'ALL' : r.toUpperCase()}
          </button>
        );
      })}
    </div>
  );
}

function SparkLine({ data = [], todayDay = 31, color = '#BC8D26' }) {
  if (!data || data.every(v => v === 0)) return null;
  const visible = data.slice(0, todayDay);
  const max = Math.max(...visible, 1);
  const W = 100, H = 32, pad = 3;
  const pts = visible.map((v, i) => {
    const x = pad + (i / Math.max(visible.length - 1, 1)) * (W - pad * 2);
    const y = H - pad - ((v / max) * (H - pad * 2));
    return `${x.toFixed(1)},${y.toFixed(1)}`;
  }).join(' ');
  const lastPt = pts.split(' ').pop() || '';
  const [lx, ly] = lastPt.split(',');
  return (
    <svg width={W} height={H} style={{ overflow: 'visible', opacity: 0.8, flexShrink: 0 }}>
      <polyline points={pts} fill="none" stroke={color} strokeWidth="1.5" strokeLinejoin="round" strokeLinecap="round" />
      {lx && <circle cx={lx} cy={ly} r="2.5" fill={color} />}
    </svg>
  );
}

function KPICard({ label, value, subtitle, target, targetValue, progress, icon: Icon, accent, sparkData, todayDay }) {
  const colors = { copper: '#BC8D26', ocean: '#5A7A99', gold: '#DBB85E', ink: '#002855' };
  const col = colors[accent] || '#BC8D26';
  return (
    <div className="premium-card p-4 relative overflow-hidden">
      <div className="absolute top-0 right-0 w-12 h-12 opacity-[0.07]" style={{ background: col, clipPath: 'polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)', transform: 'translate(30%, -30%)' }}></div>
      <div className="flex items-start justify-between mb-2">
        <p className="font-display text-[9px] tracking-[0.25em]" style={{ color: col, fontWeight: 600 }}>{label.toUpperCase()}</p>
        <Icon className="w-4 h-4 opacity-30" style={{ color: col }} />
      </div>
      <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', gap: 6 }}>
        <div style={{ minWidth: 0 }}>
          <p className="font-display text-2xl ink" style={{ fontWeight: 700 }}>{value}</p>
          {subtitle && <p className="text-[10px] italic ocean mt-0.5">{subtitle}</p>}
        </div>
        {sparkData && sparkData.some(v => v > 0) && (
          <SparkLine data={sparkData} todayDay={todayDay || 31} color={col} />
        )}
      </div>
      {target > 0 && (
        <div className="mt-2">
          <div className="flex justify-between text-[9px] mb-1">
            <span className="italic ocean">Target: {targetValue}</span>
            <span className="font-display" style={{ color: col, fontWeight: 700 }}>{progress}%</span>
          </div>
          <div style={{ height: 3, background: 'rgba(0,40,85,0.1)' }}>
            <div style={{ height: '100%', width: `${Math.min(progress, 100)}%`, background: col, transition: 'width 0.5s' }}></div>
          </div>
        </div>
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
    const sold = repVisits.filter(isSoldVisit).length;
    return { rep, monthSales, target, pct, visitCount, visitTarget, visitPct, sold };
  }).sort((a, b) => b.monthSales - a.monthSales);

  return (
    <div className="space-y-2">
      {repStats.map((s, i) => (
        <div key={s.rep} className="border border p-3 hover:border-copper/40 transition-colors">
          <div className="flex items-center gap-3 mb-2">
            <div className="w-7 h-7 flex-shrink-0 flex items-center justify-center font-display text-xs bg-ink" style={{ color: '#FCF7F2', fontWeight: 700 }}>{i + 1}</div>
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
              <div className="h-1 "><div className="h-full" style={{ width: `${s.visitPct}%`, background: '#5A7A99' }}></div></div>
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
        const outcomeColor = v.outcome === 'Sold In' ? '#5A7A99' : v.outcome === 'Rejected' ? '#CC233A' : '#BC8D26';
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
  const stages = ['New', 'Contacted', 'Converted', 'Lost'];
  const counts = stages.map(s => ({ stage: s, count: clients.filter(c => c.status === s).length }));
  const max = Math.max(...counts.map(c => c.count), 1);
  const colors = { New: '#5A7A99', Contacted: '#DBB85E', Converted: '#2d8659', Lost: '#CC233A' };
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

// =================== Annual Revenue Tracker ===================
function AnnualRevenueTracker({ visits, targets }) {
  const year = new Date().getFullYear();
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const teamMonthlyTarget = Object.values(targets).reduce((s, t) => s + (t.revenue || 0), 0);
  const annualTarget = teamMonthlyTarget * 12;

  const monthData = months.map((m, idx) => {
    const monthStr = `${year}-${String(idx + 1).padStart(2, '0')}`;
    const sales = visits
      .filter(v => (v.date || '').startsWith(monthStr))
      .reduce((s, v) => s + (v.saleAmount || 0), 0);
    return { month: m, sales, target: teamMonthlyTarget };
  });

  const totalActual = monthData.reduce((s, m) => s + m.sales, 0);
  const pct = annualTarget > 0 ? Math.min(100, Math.round((totalActual / annualTarget) * 100)) : 0;
  const currentMonth = new Date().getMonth();

  // SVG bar chart
  const barW = 28, barGap = 8, chartH = 80, pad = 4;
  const maxVal = Math.max(...monthData.map(m => Math.max(m.sales, m.target)), 1);
  const chartW = months.length * (barW + barGap) + pad * 2;

  return (
    <div>
      {/* Summary row */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: 12, marginBottom: 20 }}>
        <div style={{ textAlign: 'center', padding: '12px 8px', background: 'rgba(0,40,85,0.04)', borderLeft: '3px solid #BC8D26' }}>
          <p className="font-display" style={{ fontSize: 8, letterSpacing: '0.25em', color: '#BC8D26', fontWeight: 600 }}>ANNUAL TARGET</p>
          <p className="font-display ink" style={{ fontSize: 18, fontWeight: 700, marginTop: 2 }}>{ZAR(annualTarget)}</p>
        </div>
        <div style={{ textAlign: 'center', padding: '12px 8px', background: 'rgba(0,40,85,0.04)', borderLeft: '3px solid #5A7A99' }}>
          <p className="font-display" style={{ fontSize: 8, letterSpacing: '0.25em', color: '#5A7A99', fontWeight: 600 }}>ACTUAL YTD</p>
          <p className="font-display ink" style={{ fontSize: 18, fontWeight: 700, marginTop: 2 }}>{ZAR(totalActual)}</p>
        </div>
        <div style={{ textAlign: 'center', padding: '12px 8px', background: 'rgba(0,40,85,0.04)', borderLeft: `3px solid ${pct >= 100 ? '#2d6e4e' : '#DBB85E'}` }}>
          <p className="font-display" style={{ fontSize: 8, letterSpacing: '0.25em', color: pct >= 100 ? '#2d6e4e' : '#BC8D26', fontWeight: 600 }}>ACHIEVEMENT</p>
          <p className="font-display ink" style={{ fontSize: 18, fontWeight: 700, marginTop: 2 }}>{pct}%</p>
        </div>
      </div>

      {/* Progress bar */}
      <div style={{ marginBottom: 20 }}>
        <div style={{ height: 8, background: 'rgba(0,40,85,0.1)', borderRadius: 4 }}>
          <div style={{ height: '100%', width: `${pct}%`, background: pct >= 100 ? '#2d6e4e' : '#BC8D26', borderRadius: 4, transition: 'width 0.8s' }}></div>
        </div>
        <p style={{ fontSize: 10, color: '#5A7A99', fontStyle: 'italic', marginTop: 4 }}>
          {ZAR(annualTarget - totalActual)} remaining to annual target
        </p>
      </div>

      {/* Monthly bar chart */}
      <div style={{ overflowX: 'auto' }}>
        <svg width={chartW} height={chartH + 30} style={{ display: 'block', minWidth: '100%' }}>
          {monthData.map((m, i) => {
            const x = pad + i * (barW + barGap);
            const targetH = (m.target / maxVal) * chartH;
            const salesH = (m.sales / maxVal) * chartH;
            const isCurrentMonth = i === currentMonth;
            return (
              <g key={m.month}>
                {/* Target bar (outline) */}
                <rect x={x} y={chartH - targetH} width={barW} height={targetH}
                  fill="none" stroke="rgba(0,40,85,0.15)" strokeWidth="1" />
                {/* Actual bar */}
                <rect x={x + 2} y={chartH - salesH} width={barW - 4} height={salesH}
                  fill={isCurrentMonth ? '#BC8D26' : (m.sales >= m.target ? '#2d6e4e' : '#5A7A99')} opacity={0.8} />
                {/* Month label */}
                <text x={x + barW / 2} y={chartH + 14} textAnchor="middle"
                  style={{ fontSize: 8, fontFamily: 'Cinzel, serif', fontWeight: isCurrentMonth ? 700 : 400, fill: isCurrentMonth ? '#BC8D26' : '#002855', letterSpacing: '0.05em' }}>
                  {m.month.toUpperCase()}
                </text>
                {/* Amount label if any sales */}
                {m.sales > 0 && (
                  <text x={x + barW / 2} y={chartH - salesH - 3} textAnchor="middle"
                    style={{ fontSize: 7, fontFamily: 'Cinzel, serif', fontWeight: 600, fill: '#002855' }}>
                    {(m.sales / 1000).toFixed(0)}k
                  </text>
                )}
              </g>
            );
          })}
        </svg>
      </div>
      <div style={{ display: 'flex', gap: 16, marginTop: 8 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 10, color: '#5A7A99' }}>
          <div style={{ width: 12, height: 8, background: '#5A7A99', opacity: 0.8 }}></div> Actual
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 10, color: '#2d6e4e' }}>
          <div style={{ width: 12, height: 8, background: '#2d6e4e', opacity: 0.8 }}></div> Target met
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 10, color: '#BC8D26' }}>
          <div style={{ width: 12, height: 8, background: '#BC8D26', opacity: 0.8 }}></div> Current month
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 10, color: 'rgba(0,40,85,0.4)' }}>
          <div style={{ width: 12, height: 8, border: '1px solid rgba(0,40,85,0.3)' }}></div> Target
        </div>
      </div>
    </div>
  );
}

// =================== Monthly Close Per Agent ===================
function MonthlyCloseTable({ visits, targets }) {
  const year = new Date().getFullYear();
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const REPS = SALES_REPS;

  return (
    <div style={{ overflowX: 'auto' }}>
      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
        <thead>
          <tr style={{ background: '#002855' }}>
            <th style={{ padding: '8px 12px', textAlign: 'left', fontFamily: 'Cinzel, serif', fontSize: 9, letterSpacing: '0.2em', color: '#DBB85E', fontWeight: 600 }}>AGENT</th>
            {months.map(m => (
              <th key={m} style={{ padding: '8px 6px', textAlign: 'right', fontFamily: 'Cinzel, serif', fontSize: 8, letterSpacing: '0.1em', color: 'rgba(255,255,255,0.7)', fontWeight: 600 }}>{m.toUpperCase()}</th>
            ))}
            <th style={{ padding: '8px 12px', textAlign: 'right', fontFamily: 'Cinzel, serif', fontSize: 9, letterSpacing: '0.15em', color: '#DBB85E', fontWeight: 600 }}>TOTAL</th>
          </tr>
        </thead>
        <tbody>
          {REPS.map((rep, ri) => {
            const repTarget = targets[rep]?.revenue || 0;
            const monthSales = months.map((m, idx) => {
              const monthStr = `${year}-${String(idx + 1).padStart(2, '0')}`;
              return visits
                .filter(v => v.salesRep === rep && (v.date || '').startsWith(monthStr))
                .reduce((s, v) => s + (v.saleAmount || 0), 0);
            });
            const totalSales = monthSales.reduce((s, v) => s + v, 0);
            const currentMonthIdx = new Date().getMonth();

            return (
              <tr key={rep} style={{ background: ri % 2 === 0 ? '#FCF7F2' : 'rgba(0,40,85,0.03)', borderBottom: '1px solid rgba(0,40,85,0.1)' }}>
                <td style={{ padding: '10px 12px', fontFamily: 'Cinzel, serif', fontWeight: 700, color: '#002855', fontSize: 12, letterSpacing: '0.08em' }}>
                  {rep.toUpperCase()}
                  {repTarget > 0 && <div style={{ fontSize: 9, color: '#5A7A99', fontWeight: 400, fontStyle: 'italic', marginTop: 1 }}>Target: {ZAR(repTarget)}/mo</div>}
                </td>
                {monthSales.map((sales, idx) => {
                  const isCurrentMonth = idx === currentMonthIdx;
                  const metTarget = repTarget > 0 && sales >= repTarget;
                  return (
                    <td key={idx} style={{
                      padding: '10px 6px', textAlign: 'right', fontSize: 11,
                      fontFamily: 'Cinzel, serif', fontWeight: sales > 0 ? 700 : 400,
                      color: sales === 0 ? 'rgba(0,40,85,0.25)' : metTarget ? '#2d6e4e' : isCurrentMonth ? '#BC8D26' : '#002855',
                      background: isCurrentMonth ? 'rgba(188,141,38,0.06)' : 'transparent',
                    }}>
                      {sales === 0 ? '—' : `${(sales / 1000).toFixed(0)}k`}
                    </td>
                  );
                })}
                <td style={{ padding: '10px 12px', textAlign: 'right', fontFamily: 'Cinzel, serif', fontWeight: 700, color: '#BC8D26', fontSize: 12 }}>
                  {ZAR(totalSales)}
                </td>
              </tr>
            );
          })}
          {/* Team total row */}
          <tr style={{ background: '#002855', borderTop: '2px solid #BC8D26' }}>
            <td style={{ padding: '10px 12px', fontFamily: 'Cinzel, serif', fontWeight: 700, color: '#DBB85E', fontSize: 11, letterSpacing: '0.15em' }}>TEAM TOTAL</td>
            {months.map((m, idx) => {
              const monthStr = `${year}-${String(idx + 1).padStart(2, '0')}`;
              const total = visits
                .filter(v => (v.date || '').startsWith(monthStr))
                .reduce((s, v) => s + (v.saleAmount || 0), 0);
              const isCurrentMonth = idx === new Date().getMonth();
              return (
                <td key={idx} style={{
                  padding: '10px 6px', textAlign: 'right', fontSize: 11,
                  fontFamily: 'Cinzel, serif', fontWeight: total > 0 ? 700 : 400,
                  color: total === 0 ? 'rgba(255,255,255,0.3)' : isCurrentMonth ? '#DBB85E' : '#FCF7F2',
                  background: isCurrentMonth ? 'rgba(219,184,94,0.1)' : 'transparent',
                }}>
                  {total === 0 ? '—' : `${(total / 1000).toFixed(0)}k`}
                </td>
              );
            })}
            <td style={{ padding: '10px 12px', textAlign: 'right', fontFamily: 'Cinzel, serif', fontWeight: 700, color: '#DBB85E', fontSize: 12 }}>
              {ZAR(visits.reduce((s, v) => s + (v.saleAmount || 0), 0))}
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

// =================== MANAGER PORTAL ===================
function ManagerPortal({ targets, saveTargets, clients, visits, askConfirm, skuPrices, saveSkuPrices, showToast }) {
  const [draft, setDraft] = useState(targets);
  const [skuDraftPrices, setSkuDraftPrices] = useState({});
  const [savedFlash, setSavedFlash] = useState(false);

  useEffect(() => { setDraft(targets); }, [targets]);
  useEffect(() => { if (skuPrices) setSkuDraftPrices(skuPrices); }, [skuPrices]);

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
    // Build client channel lookup map for backfilling empty visit channels
    const clientChannelMap = {};
    clients.forEach(c => { clientChannelMap[Number(c.id)] = c.channel; });

    SALES_REPS.forEach(rep => {
      const repVisits = visits.filter(v => v.salesRep === rep && v.date && v.date.startsWith(m));
      // Get effective channel — use visit channel if set, otherwise look up from client
      const getChannel = (v) => v.channel || clientChannelMap[Number(v.clientId)] || '';
      out[rep] = {
        revenue:     repVisits.reduce((s, v) => s + (v.saleAmount || 0), 0),
        visits:      repVisits.length,
        tradeRetail: repVisits.filter(v => getChannel(v) === 'Trade').length,
        onCon:       repVisits.filter(v => getChannel(v) === 'On-Con').length,
        b2b:         repVisits.filter(v => getChannel(v) === 'B2B').length,
      };
    });
    return out;
  }, [visits, clients, m]);

  // Team totals (from draft)
  const teamTotals = useMemo(() => SALES_REPS.reduce((acc, r) => ({
    revenue:     acc.revenue     + (draft[r]?.revenue     || 0),
    visits:      acc.visits      + (draft[r]?.visits      || 0),
    tradeRetail: acc.tradeRetail + (draft[r]?.tradeRetail || 0),
    onCon:       acc.onCon       + (draft[r]?.onCon       || 0),
    b2b:         acc.b2b         + (draft[r]?.b2b         || 0),
  }), { revenue: 0, visits: 0, tradeRetail: 0, onCon: 0, b2b: 0 }), [draft]);

  // Client distribution per rep per channel
  const repBook = useMemo(() => {
    const out = {};
    SALES_REPS.forEach(rep => {
      const own = clients.filter(c => c.accountManager === rep);
      out[rep] = {
        total: own.length,
        b2b: own.filter(c => c.channel === 'B2B').length,
        tradeRetail: own.filter(c => c.channel === 'Trade').length,
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

    // Sheet 3b: Client Order History (one row per order, with full SKU detail)
    const orderHistoryRows = [];
    visits
      .filter(v => v.items && v.items.length > 0)
      .sort((a, b) => (b.date || '').localeCompare(a.date || ''))
      .forEach(v => {
        const c = clients.find(cl => cl.id === v.clientId) || {};
        const skuSummary = v.items.map(it => `${it.qty}x ${it.name}${Number(it.unitPrice) < Number(it.listPrice) ? ` (disc. ${ZAR(it.unitPrice)})` : ''}`).join(' | ');
        orderHistoryRows.push({
          'Date': v.date,
          'Venue': c.venue || 'Unknown',
          'Channel': c.channel || '',
          'Location': c.location || '',
          'Account Manager': c.accountManager || '',
          'Sales Rep': v.salesRep || '',
          'Outcome': v.outcome || '',
          'SKUs Ordered': skuSummary,
          'Total SKU Lines': v.items.length,
          'Total Units': v.items.reduce((s, it) => s + (Number(it.qty) || 0), 0),
          'Order Total Ex VAT (R)': Number(v.saleAmount || 0).toFixed(2),
          'Visit Notes': v.notes || '',
          'Follow-up Notes': v.followUp || '',
        });
      });
    const wsOrders = XLSX.utils.json_to_sheet(orderHistoryRows.length > 0 ? orderHistoryRows : [{ 'Note': 'No orders placed yet.' }]);
    wsOrders['!cols'] = [
      { wch: 12 }, { wch: 28 }, { wch: 14 }, { wch: 18 }, { wch: 14 },
      { wch: 12 }, { wch: 16 }, { wch: 60 }, { wch: 14 }, { wch: 12 }, { wch: 20 }, { wch: 40 }, { wch: 40 },
    ];
    XLSX.utils.book_append_sheet(wb, wsOrders, 'Client Order History');


    const targetRows = SALES_REPS.map(rep => ({
      'Sales Rep': rep,
      'Revenue Target (R)': Number(targets[rep]?.revenue || 0).toFixed(2),
      'Visits Target': targets[rep]?.visits || 0,
      'Trade Visits Target': targets[rep]?.tradeRetail || 0,
      'On-Con Visits Target': targets[rep]?.onCon || 0,
      'B2B Visits Target': targets[rep]?.b2b || 0,
    }));
    targetRows.push({
      'Sales Rep': 'TEAM TOTAL',
      'Revenue Target (R)': SALES_REPS.reduce((s, r) => s + (targets[r]?.revenue || 0), 0).toFixed(2),
      'Visits Target': SALES_REPS.reduce((s, r) => s + (targets[r]?.visits || 0), 0),
      'Trade Visits Target': SALES_REPS.reduce((s, r) => s + (targets[r]?.tradeRetail || 0), 0),
      'On-Con Visits Target': SALES_REPS.reduce((s, r) => s + (targets[r]?.onCon || 0), 0),
      'B2B Visits Target': SALES_REPS.reduce((s, r) => s + (targets[r]?.b2b || 0), 0),
    });
    const wsTargets = XLSX.utils.json_to_sheet(targetRows);
    wsTargets['!cols'] = [{ wch: 16 }, { wch: 20 }, { wch: 18 }, { wch: 18 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, wsTargets, 'Targets');

    // Sheet 5: Performance vs Target (current month)
    const monthLabel = new Date().toLocaleDateString('en-ZA', { month: 'long', year: 'numeric' });
    const m = monthISO();
    // Client channel lookup so visits with an empty channel still classify correctly
    const exportClientChannelMap = {};
    clients.forEach(c => { exportClientChannelMap[Number(c.id)] = c.channel; });
    const getChannel = (v) => v.channel || exportClientChannelMap[Number(v.clientId)] || '';

    const perfRows = SALES_REPS.map(rep => {
      const repVisits = visits.filter(v => v.salesRep === rep && v.date && v.date.startsWith(m));
      const revenue = repVisits.reduce((s, v) => s + (v.saleAmount || 0), 0);
      const visitCount = repVisits.length;
      const tradeCount = repVisits.filter(v => getChannel(v) === 'Trade').length;
      const onConCount = repVisits.filter(v => getChannel(v) === 'On-Con').length;
      const b2bCount = repVisits.filter(v => getChannel(v) === 'B2B').length;
      const sold = repVisits.filter(isSoldVisit).length;
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
        'Trade Visits': tradeCount,
        'Trade Target': t.tradeRetail || 0,
        'On-Con Visits': onConCount,
        'On-Con Target': t.onCon || 0,
        'B2B Visits': b2bCount,
        'B2B Target': t.b2b || 0,
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
      <div className="relative overflow-hidden border-2 border-ink" style={{ background: '#002855' }}>
        <RaysBackdrop opacity={0.1} />
        <div className="relative p-4 md:p-10">
          <div className="flex items-center gap-3 mb-3">
            <AvanteLogo height={44} />
            <div>
              <p className="font-display text-[9px] tracking-[0.3em]" style={{ color: '#DBB85E', fontWeight: 600 }}>STRATEGIC EXECUTION</p>
              <h1 className="font-display text-xl md:text-4xl" style={{ color: '#FCF7F2', fontWeight: 700, letterSpacing: '0.06em' }}>MANAGER PORTAL</h1>
            </div>
          </div>
          {/* Action buttons — row on desktop, 2-col grid on mobile */}
          <div className="grid grid-cols-2 md:flex md:flex-row gap-2 mt-2">
            {savedFlash && (
              <span className="col-span-2 inline-flex items-center justify-center gap-2 px-3 py-2 font-display text-[10px] tracking-[0.2em]" style={{ color: '#002855', background: '#DBB85E', fontWeight: 700 }}>
                <Save className="w-3.5 h-3.5" /> SAVED
              </span>
            )}
<button onClick={handleSave} disabled={!dirty} className={`col-span-2 md:col-span-1 flex items-center justify-center gap-2 px-5 py-2.5 font-display text-xs tracking-[0.2em] transition-colors ${dirty ? 'bg-copper' : 'bg-copper/40 cursor-not-allowed'}`} style={{ color: '#FCF7F2', fontWeight: 700 }}>
              <Save className="w-4 h-4" /> SAVE TARGETS
            </button>
          </div>
        </div>
      </div>

      {/* Data Export */}
      <div className="premium-card p-4 md:p-6 relative overflow-hidden">
        <div className="flex items-start gap-3 mb-3">
          <div className="w-10 h-10 flex-shrink-0 flex items-center justify-center bg-ink">
            <FileSpreadsheet className="w-5 h-5" style={{ color: '#DBB85E' }} />
          </div>
          <div className="flex-1 min-w-0">
            <p className="font-display text-[9px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>DATA EXPORT</p>
            <h2 className="font-display text-lg ink" style={{ fontWeight: 700 }}>EXPORT TO EXCEL</h2>
            <p className="italic text-xs ocean mt-0.5">All clients, visits, SKUs, targets &amp; performance in one workbook.</p>
          </div>
        </div>
        <button onClick={exportToExcel} className="w-full flex items-center justify-center gap-2 bg-ink px-4 py-3 font-display text-xs tracking-[0.25em]" style={{ color: '#FCF7F2', fontWeight: 700 }}>
          <Download className="w-4 h-4" /> DOWNLOAD .XLSX
        </button>
        <p style={{ fontSize: 10, color: '#5A7A99', fontStyle: 'italic', marginTop: 6, textAlign: 'center' }}>
          Download weekly and email to matthew@breakfreebeverages.com
        </p>

      </div>

      {/* Team aggregate cards */}
      <div>
        <div className="flex items-center gap-3 mb-3">
          <div className="w-2 h-2 bg-copper diamond-clip"></div>
          <h2 className="font-display text-sm tracking-[0.3em] ink" style={{ fontWeight: 700 }}>TEAM AGGREGATE TARGETS</h2>
        </div>
        <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
          <AggregateCard label="Total Revenue Target" value={ZAR(teamTotals.revenue)} icon={DollarSign} color="#BC8D26" />
          <AggregateCard label="Total Visits Target"  value={teamTotals.visits}        icon={Activity}   color="#5A7A99" />
          <AggregateCard label="Trade Visits"   value={teamTotals.tradeRetail}   icon={Briefcase}  color="#002855" />
          <AggregateCard label="On-Con Visits"     value={teamTotals.onCon}         icon={Activity}   color="#5A7A99" />
          <AggregateCard label="B2B Visits"        value={teamTotals.b2b}           icon={TrendingUp} color="#BC8D26" />
        </div>
      </div>

      {/* SKU Price Editor */}
      <div className="premium-card p-4 md:p-6">
        <div className="flex items-center justify-between mb-4">
          <div className="flex items-center gap-3">
            <div className="w-2 h-2 bg-copper diamond-clip"></div>
            <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>SKU PRICE MANAGEMENT</h2>
          </div>
          <p className="italic ocean" style={{ fontSize: 10 }}>Prices update immediately in Log Visit</p>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 12 }}>
          {SKU_CATALOGUE.map(sku => (
            <div key={sku.id} style={{ padding: '10px 12px', border: '1px solid rgba(0,40,85,0.15)', background: '#FCF7F2' }}>
              <p className="font-display" style={{ fontSize: 9, letterSpacing: '0.2em', color: '#BC8D26', fontWeight: 600, marginBottom: 4 }}>{sku.name.toUpperCase()}</p>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <span className="font-display ink" style={{ fontSize: 12, fontWeight: 700, flexShrink: 0 }}>R</span>
                <input
                  type="number"
                  min="0"
                  step="0.01"
                  value={skuDraftPrices[sku.id] ?? sku.price}
                  onChange={e => setSkuDraftPrices(prev => ({ ...prev, [sku.id]: parseFloat(e.target.value) || 0 }))}
                  style={{ flex: 1, padding: '6px 8px', border: '1px solid rgba(0,40,85,0.2)', background: '#FCF7F2', fontFamily: "'Cinzel', serif", fontSize: 13, fontWeight: 700, color: '#002855', outline: 'none', width: '100%' }}
                />
              </div>
            </div>
          ))}
        </div>
        <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 12, gap: 8 }}>
          <button onClick={() => setSkuDraftPrices({})}
            style={{ padding: '8px 16px', border: '1px solid rgba(0,40,85,0.2)', background: 'none', fontFamily: "'Cinzel', serif", fontSize: 9, letterSpacing: '0.2em', color: '#002855', cursor: 'pointer', fontWeight: 600 }}>
            RESET TO DEFAULT
          </button>
          <button onClick={() => { saveSkuPrices(skuDraftPrices); showToastMsg('SKU prices saved!'); }}
            style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '8px 20px', background: '#BC8D26', border: 'none', color: '#FCF7F2', fontFamily: "'Cinzel', serif", fontSize: 10, letterSpacing: '0.2em', fontWeight: 700, cursor: 'pointer' }}>
            <Save style={{ width: 14, height: 14 }} /> SAVE PRICES
          </button>
        </div>
      </div>

      {/* ── ANNUAL REVENUE TRACKER ── */}
      <div className="premium-card p-4 md:p-6">
        <div className="flex items-center justify-between mb-4 flex-wrap gap-3">
          <div className="flex items-center gap-3">
            <div className="w-2 h-2 bg-copper diamond-clip"></div>
            <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>ANNUAL REVENUE TRACKER</h2>
          </div>
          <p className="italic ocean" style={{ fontSize: 10 }}>
            {new Date().getFullYear()} — Cumulative team sales
          </p>
        </div>
        <AnnualRevenueTracker visits={visits} targets={targets} />
      </div>

      {/* ── PER-AGENT MONTHLY CLOSE ── */}
      <div className="premium-card p-4 md:p-6">
        <div className="flex items-center gap-3 mb-4">
          <div className="w-2 h-2 bg-copper diamond-clip"></div>
          <h2 className="font-display text-xs md:text-sm tracking-[0.2em] ink" style={{ fontWeight: 700 }}>MONTHLY CLOSE — PER AGENT</h2>
        </div>
        <MonthlyCloseTable visits={visits} targets={targets} />
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
              onSave={async () => {
                // Build the full draft with latest changes and save
                await saveTargets({ ...draft });
                setSavedFlash(true);
                setTimeout(() => setSavedFlash(false), 2000);
              }}
            />
          ))}
        </div>
      </div>

      {/* Save reminder bar — sits above mobile bottom nav */}
      {dirty && (
        <div className="sticky bottom-20 md:bottom-4 z-30 max-w-2xl mx-auto">
          <div className="flex items-center justify-between gap-3 p-3 md:p-4 border-2 border-copper shadow-2xl" style={{ background: '#002855' }}>
            <div className="flex items-center gap-2">
              <div className="w-2 h-2 bg-gold diamond-clip animate-pulse flex-shrink-0"></div>
              <p className="font-display text-[9px] md:text-xs tracking-[0.15em]" style={{ color: '#FCF7F2', fontWeight: 600 }}>UNSAVED CHANGES</p>
            </div>
            <div className="flex items-center gap-2">
              <button onClick={handleSave} className="flex items-center gap-1.5 bg-copper px-3 md:px-4 py-2 font-display text-[9px] md:text-[10px] tracking-[0.2em]" style={{ color: '#FCF7F2', fontWeight: 700 }}>
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

function RepTargetCard({ rep, draft, perf, book, onChange, onSave }) {
  const [saved, setSaved] = useState(false);

  const handleSave = async () => {
    await onSave();
    setSaved(true);
    setTimeout(() => setSaved(false), 2000);
  };
  // Total Visits is always the sum of Trade + On-Con + B2B — not editable
  const computedVisits = (draft?.tradeRetail || 0) + (draft?.onCon || 0) + (draft?.b2b || 0);

  const fields = [
    { key: 'revenue',     label: 'Revenue Target', icon: DollarSign, prefix: 'R', perfKey: 'revenue',     isCurrency: true, hint: 'Monthly Rand target' },
    { key: 'tradeRetail', label: 'Trade',          icon: Briefcase,  perfKey: 'tradeRetail', hint: 'Trade visits KPI' },
    { key: 'onCon',       label: 'On-Con',         icon: Activity,   perfKey: 'onCon',       hint: 'On-Con visits KPI' },
    { key: 'b2b',         label: 'B2B',            icon: TrendingUp, perfKey: 'b2b',         hint: 'B2B visits KPI' },
  ];

  return (
    <div className="premium-card p-4 md:p-6 relative">
      <div className="flex items-center justify-between pb-3 mb-3 border-b">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 flex items-center justify-center bg-ink">
            <span className="font-display text-sm" style={{ color: '#DBB85E', fontWeight: 700 }}>{rep[0]}</span>
          </div>
          <div>
            <h3 className="font-display text-base tracking-[0.1em] ink" style={{ fontWeight: 700 }}>{rep.toUpperCase()}</h3>
            <p className="text-[10px] italic ocean">{book.total} clients</p>
          </div>
        </div>
        <Target className="w-4 h-4 copper" />
      </div>

      <div className="space-y-3">
        {/* Revenue + Private + Trade editable fields */}
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
                  onChange={(e) => {
                    onChange(f.key, e.target.value);
                    // Auto-update visits total whenever any channel changes
                    if (['tradeRetail','onCon','b2b'].includes(f.key)) {
                      const newTrd = f.key === 'tradeRetail' ? Number(e.target.value)||0 : (draft?.tradeRetail||0);
                      const newOnC = f.key === 'onCon'        ? Number(e.target.value)||0 : (draft?.onCon||0);
                      const newB2b = f.key === 'b2b'          ? Number(e.target.value)||0 : (draft?.b2b||0);
                      onChange('visits', newTrd + newOnC + newB2b);
                    }
                  }}
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

        {/* Total Visits — read-only, auto-sum */}
        {(() => {
          const actual = perf?.visits || 0;
          const pct = computedVisits > 0 ? Math.min(100, Math.round((actual / computedVisits) * 100)) : 0;
          return (
            <div style={{ background: 'rgba(0,40,85,0.04)', padding: '10px 12px', borderLeft: '2px solid rgba(0,40,85,0.2)' }}>
              <div className="flex items-center gap-1.5 mb-1">
                <Activity className="w-3 h-3 copper" />
                <label className="font-display text-[9px] tracking-[0.2em] copper" style={{ fontWeight: 600 }}>TOTAL VISITS</label>
                <span style={{ fontSize: 8, fontStyle: 'italic', color: '#5A7A99', marginLeft: 4 }}>auto-calculated</span>
              </div>
              <div className="flex items-center gap-2">
                <span className="font-display ink text-sm flex-1" style={{ fontWeight: 700, fontSize: 18 }}>{computedVisits}</span>
                <span style={{ fontSize: 10, color: '#5A7A99', fontStyle: 'italic' }}>= Trade + On-Con + B2B</span>
              </div>
              <div className="flex items-center justify-between mt-1.5 text-[10px]">
                <span className="ocean italic">Total visits · actual: <span className="ink font-display" style={{ fontWeight: 700 }}>{actual}</span></span>
                <span className="font-display tracking-wider copper flex-shrink-0" style={{ fontWeight: 700 }}>{pct}%</span>
              </div>
              <div className="h-1 mt-1">
                <div className="h-full bg-copper transition-all duration-500" style={{ width: `${pct}%` }}></div>
              </div>
            </div>
          );
        })()}

        {/* Per-rep save button */}
        <button
          onClick={handleSave}
          style={{ width: '100%', marginTop: 12, padding: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: saved ? '#2d8659' : '#002855', border: 'none', fontFamily: '"Cinzel",serif', fontSize: 9, letterSpacing: '0.25em', fontWeight: 700, color: '#FCF7F2', cursor: 'pointer', transition: 'background 0.2s' }}>
          <Save style={{ width: 12, height: 12 }} />
          {saved ? 'SAVED ✓' : `SAVE ${rep.toUpperCase()}'S TARGETS`}
        </button>
      </div>
    </div>
  );
}

// =================== Leads Page ===================
function KanbanView({ filtered, onSelect, onDelete, updateClient, activeChannel }) {
  // Determine which status columns to show — if filtering by B2B, show B2B stages; otherwise standard
  const PRIORITY_ORDER = { 'High': 0, 'Medium': 1, 'Low': 2 };
  const statusCols = activeChannel === 'B2B' ? B2B_STATUS_LABELS : STATUSES;

  return (
    <div style={{ display: 'grid', gridTemplateColumns: `repeat(${statusCols.length}, minmax(180px, 1fr))`, gap: 12, overflowX: 'auto', minWidth: 0 }}>
      {statusCols.map(status => {
        const color = getStatusColor(status);
        const pct = activeChannel === 'B2B' ? getB2BPct(status) : null;
        const cols = filtered
          .filter(c => c.status === status)
          .sort((a, b) => (PRIORITY_ORDER[a.priority] ?? 1) - (PRIORITY_ORDER[b.priority] ?? 1));
        return (
          <div key={status} style={{ background: `${color}12`, border: `1px solid ${color}33`, minWidth: 180 }}>
            {/* Column header */}
            <div style={{ padding: '10px 12px', borderBottom: `2px solid ${color}`, display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: color }}>
              <span style={{ fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.2em', fontWeight: 700, color: '#FCF7F2' }}>{status.toUpperCase()}</span>
              <span style={{ fontFamily: "'Cinzel',serif", fontSize: 9, fontWeight: 700, color: 'rgba(255,255,255,0.8)', background: 'rgba(255,255,255,0.2)', borderRadius: 10, padding: '1px 8px' }}>{cols.length}</span>
            </div>
            {/* Cards — sorted High → Medium → Low */}
            <div style={{ padding: 8, display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 560, overflowY: 'auto' }}>
              {cols.length === 0 && (
                <p style={{ fontSize: 11, color: 'rgba(0,40,85,0.3)', fontStyle: 'italic', textAlign: 'center', padding: '16px 0' }}>No clients</p>
              )}
              {cols.map(c => (
                <div key={c.id}
                  onClick={() => onSelect(c)}
                  style={{ background: '#FCF7F2', border: '1px solid rgba(0,40,85,0.12)', borderTop: `3px solid ${c.priority === 'High' ? '#CC233A' : c.priority === 'Medium' ? '#BC8D26' : '#5A7A99'}`, padding: '10px 12px', cursor: 'pointer', transition: 'box-shadow 0.15s' }}
                  onMouseEnter={e => e.currentTarget.style.boxShadow = '0 2px 12px rgba(0,40,85,0.12)'}
                  onMouseLeave={e => e.currentTarget.style.boxShadow = 'none'}>
                  <p style={{ fontFamily: "'Cinzel',serif", fontWeight: 700, fontSize: 11, color: '#002855', marginBottom: 4, lineHeight: 1.3 }}>{c.venue}</p>
                  <p style={{ fontSize: 10, color: '#5A7A99', fontStyle: 'italic', marginBottom: 6 }}>
                    {c.accountManager}{c.location ? ` · ${c.location}` : ''}
                  </p>
                  {/* Client Tags */}
                  {c.clientTags && c.clientTags.length > 0 && (
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 3, marginBottom: 6 }}>
                      {c.clientTags.map(tag => (
                        <span key={tag} style={{ padding: '2px 6px', background: 'rgba(90,122,153,0.12)', color: '#5A7A99', fontFamily: "'Cinzel',serif", fontSize: 7, letterSpacing: '0.1em', fontWeight: 700 }}>@{tag}</span>
                      ))}
                    </div>
                  )}
                  {/* Tags */}
                  {c.tags && c.tags.length > 0 && (
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 3, marginBottom: 6 }}>
                      {c.tags.map(tag => (
                        <span key={tag} style={{ padding: '2px 6px', background: 'rgba(188,141,38,0.15)', color: '#BC8D26', fontFamily: "'Cinzel',serif", fontSize: 7, letterSpacing: '0.1em', fontWeight: 700 }}>{tag}</span>
                      ))}
                    </div>
                  )}
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginTop: 4 }}>
                    <span style={{ fontSize: 9, color: '#BC8D26', fontFamily: "'Cinzel',serif", fontWeight: 600 }}>{c.channel}</span>
                    {c.priority && <span style={{ fontSize: 8, color: 'rgba(0,40,85,0.4)', fontStyle: 'italic' }}>{c.priority}</span>}
                  </div>
                </div>
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function LeadsPage({ clients, visits, updateClient, onSelect, onAddNew, onDelete, onNavigate, defaultRepFilter = 'All', lockRepFilter = false }) {
  const [search, setSearch] = useState('');
  const [filterRep, setFilterRep] = useState(defaultRepFilter);
  const [filterChannel, setFilterChannel] = useState('All');
  const [filterLocation, setFilterLocation] = useState('All');
  const [filterStatus, setFilterStatus] = useState('All');
  const [viewMode, setViewMode] = useState('list'); // 'list' | 'kanban'
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);

  // Alphabetically sorted + filtered list
  const filtered = useMemo(() => {
    return clients
      .filter(c => {
        if (filterRep !== 'All' && c.accountManager !== filterRep) return false;
        if (filterChannel !== 'All' && c.channel !== filterChannel) return false;
        if (filterLocation !== 'All' && c.location !== filterLocation) return false;
        if (filterStatus !== 'All' && c.status !== filterStatus) return false;
        if (search.trim()) {
          const q = search.trim().toLowerCase();
          const hay = [c.venue, c.firstName, c.lastName, c.location, c.email, c.phone, c.notes]
            .filter(Boolean).join(' ').toLowerCase();
          if (!hay.includes(q)) return false;
        }
        return true;
      })
      .sort((a, b) => (a.venue || '').localeCompare(b.venue || ''));
  }, [clients, filterRep, filterChannel, filterLocation, filterStatus, search]);

  return (
    <div className="space-y-4 fade-up">
      {/* Overdue Follow-ups Kanban */}
      <OverdueClients
        clients={clients}
        visits={visits}
        visibleReps={lockRepFilter ? [defaultRepFilter] : SALES_REPS}
        onNavigate={(clientId) => { if (onSelect && clientId) { const c = clients.find(cl => cl.id === clientId); if (c) onSelect(c); } }}
      />

      {/* Page header */}
      <div className="pb-3 border-b">
        <p className="font-display text-[9px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>CLIENT DATABASE</p>
        <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap', marginTop: 4 }}>
          <div>
            <h1 className="font-display mt-1 ink" style={{ fontWeight: 700, fontSize: 28 }}>LEADS &amp; CLIENTS</h1>
            <p className="italic ocean" style={{ fontSize: 12, marginTop: 2 }}>{clients.length} venues tracked</p>
          </div>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
            {/* List / Kanban toggle */}
            <div style={{ display: 'flex', border: '1px solid rgba(0,40,85,0.2)', overflow: 'hidden' }}>
              {[['list','≡ LIST'], ['kanban','⊞ KANBAN']].map(([mode, label]) => (
                <button key={mode} onClick={() => setViewMode(mode)}
                  style={{ padding: '8px 14px', fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.15em', fontWeight: 700, border: 'none', background: viewMode === mode ? '#002855' : 'transparent', color: viewMode === mode ? '#FCF7F2' : '#002855', cursor: 'pointer' }}>
                  {label}
                </button>
              ))}
            </div>
            <button onClick={() => onAddNew && onAddNew()}
              style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '10px 18px', background: '#BC8D26', border: 'none', color: '#FCF7F2', fontFamily: "'Cinzel', serif", fontSize: 10, letterSpacing: '0.2em', fontWeight: 700, cursor: 'pointer', flexShrink: 0 }}>
              <span style={{ fontSize: 16, fontWeight: 300 }}>+</span> ADD NEW CLIENT
            </button>
          </div>
        </div>
      </div>

      {/* Search + filters */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, border: '1px solid rgba(0,40,85,0.25)', background: '#FCF7F2', padding: '0 12px' }}>
          <Search style={{ width: 16, height: 16, color: '#5A7A99', flexShrink: 0 }} />
          <input type="text" placeholder="Search venue, contact, location..."
            value={search} onChange={(e) => setSearch(e.target.value)}
            style={{ flex: 1, border: 'none', background: 'transparent', padding: '12px 0', fontFamily: "'Libre Baskerville', Georgia, serif", fontSize: 14, color: '#002855', outline: 'none' }} />
          {search && <button onClick={() => setSearch('')} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#5A7A99', fontSize: 18, lineHeight: 1 }}>×</button>}
        </div>
        {/* 4-col filter row */}
        <div style={{ display: 'grid', gridTemplateColumns: lockRepFilter ? '1fr 1fr 1fr' : 'repeat(4,1fr)', gap: 8 }}>
          {!lockRepFilter && <FilterSelect label="Rep" value={filterRep} onChange={setFilterRep} options={['All', ...SALES_REPS, 'Unassigned']} />}
          <FilterSelect label="Channel" value={filterChannel} onChange={setFilterChannel} options={['All', ...CHANNELS]} />
          <FilterSelect label="Location" value={filterLocation} onChange={setFilterLocation} options={['All', ...LOCATIONS]} />
          <FilterSelect label="Status" value={filterStatus} onChange={setFilterStatus} options={['All', ...STATUSES, ...B2B_STATUS_LABELS.filter(s => !STATUSES.includes(s))]} />
        </div>
        <p style={{ fontSize: 11, fontStyle: 'italic', color: '#5A7A99', margin: 0 }}>
          {filtered.length} of {clients.length} clients{search ? ` matching "${search}"` : ''}
        </p>
      </div>

      {/* Kanban or List view */}
      {viewMode === 'kanban' ? (
        <KanbanView filtered={filtered} onSelect={onSelect} onDelete={onDelete} updateClient={updateClient} activeChannel={filterChannel} />
      ) : (
      <div className="premium-card" style={{ overflow: 'hidden' }}>
        {filtered.length === 0 ? (
          <div style={{ padding: 40, textAlign: 'center' }}>
            <p style={{ fontStyle: 'italic', color: '#5A7A99', fontSize: 14 }}>
              {search ? `No clients found matching "${search}"` : 'No clients match your filters.'}
            </p>
            {search && (
              <button onClick={() => setSearch('')}
                style={{ marginTop: 12, padding: '8px 20px', background: '#BC8D26', color: '#FCF7F2', border: 'none', cursor: 'pointer', fontFamily: "'Cinzel', serif", fontSize: 10, letterSpacing: '0.2em', fontWeight: 700 }}>
                CLEAR SEARCH
              </button>
            )}
          </div>
        ) : (
          <div style={{ maxHeight: 640, overflowY: 'auto' }}>
            {filtered.map((c, i) => (
              <div key={c.id}>
                {/* Confirm delete inline banner */}
                {confirmDeleteId === c.id && (
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 16px', background: '#fdf0f0', borderTop: i === 0 ? 'none' : '1px solid rgba(204,35,58,0.2)', borderLeft: '3px solid #CC233A' }}>
                    <p style={{ fontSize: 12, color: '#CC233A', fontStyle: 'italic', margin: 0 }}>Remove <strong>{c.venue}</strong>? This cannot be undone.</p>
                    <div style={{ display: 'flex', gap: 8, flexShrink: 0 }}>
                      <button onClick={(e) => { e.stopPropagation(); setConfirmDeleteId(null); }}
                        style={{ padding: '5px 12px', border: '1px solid rgba(0,40,85,0.2)', background: 'none', fontFamily: "'Cinzel', serif", fontSize: 9, letterSpacing: '0.15em', color: '#002855', cursor: 'pointer', fontWeight: 600 }}>
                        CANCEL
                      </button>
                      <button onClick={(e) => { e.stopPropagation(); setConfirmDeleteId(null); onDelete && onDelete(c.id); }}
                        style={{ padding: '5px 12px', background: '#CC233A', border: 'none', fontFamily: "'Cinzel', serif", fontSize: 9, letterSpacing: '0.15em', color: '#FCF7F2', cursor: 'pointer', fontWeight: 700 }}>
                        DELETE
                      </button>
                    </div>
                  </div>
                )}
                <div onClick={() => onSelect(c)}
                  style={{
                    display: 'flex', alignItems: 'center', gap: 12, padding: '12px 16px',
                    borderTop: i === 0 ? 'none' : '1px solid rgba(0,40,85,0.1)',
                    borderLeft: `3px solid ${c.channel === 'B2B' ? '#BC8D26' : '#5A7A99'}`,
                    cursor: 'pointer', transition: 'background 0.12s',
                    background: '#FCF7F2',
                  }}
                  onMouseEnter={e => e.currentTarget.style.background = 'rgba(0,40,85,0.04)'}
                  onMouseLeave={e => e.currentTarget.style.background = '#FCF7F2'}
                >
                  {/* Venue + notes */}
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                      <span className="font-display ink" style={{ fontWeight: 700, fontSize: 13 }}>{c.venue}</span>
                      <StatusBadge status={c.status} channel={c.channel} />
                    </div>
                    <p style={{ fontSize: 11, color: '#5A7A99', fontStyle: 'italic', margin: '2px 0 0', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {[c.firstName, c.lastName].filter(Boolean).join(' ')}
                      {c.location ? ((c.firstName || c.lastName) ? ' · ' : '') + c.location : ''}
                    </p>
                  </div>
                  {/* Channel badge */}
                  <span style={{ fontSize: 9, fontFamily: "'Cinzel', serif", fontWeight: 600, letterSpacing: '0.15em', padding: '2px 6px', border: `1px solid ${c.channel === 'B2B' ? '#002855' : '#5A7A99'}`, color: c.channel === 'B2B' ? '#002855' : '#5A7A99', flexShrink: 0 }}>
                    {c.channel === 'Trade' ? 'TRADE' : c.channel === 'On-Con' ? 'ON-CON' : c.channel === 'B2B' ? 'B2B' : '—'}
                  </span>
                  {/* Rep */}
                  <span className="font-display" style={{ fontSize: 11, fontWeight: 700, color: '#002855', flexShrink: 0, minWidth: 48, textAlign: 'right' }}>
                    {c.accountManager === 'Unassigned' ? '—' : c.accountManager}
                  </span>
                  {/* Sales */}
                  {c.totalSales > 0 && (
                    <span className="font-display copper" style={{ fontSize: 11, fontWeight: 700, flexShrink: 0 }}>{ZAR(c.totalSales)}</span>
                  )}
                  {/* Delete button */}
                  {onDelete && (
                    <button
                      type="button"
                      onClick={(e) => { e.stopPropagation(); setConfirmDeleteId(confirmDeleteId === c.id ? null : c.id); }}
                      style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '4px', color: confirmDeleteId === c.id ? '#CC233A' : 'rgba(0,40,85,0.25)', flexShrink: 0, transition: 'color 0.15s' }}
                      onMouseEnter={e => e.currentTarget.style.color = '#CC233A'}
                      onMouseLeave={e => e.currentTarget.style.color = confirmDeleteId === c.id ? '#CC233A' : 'rgba(0,40,85,0.25)'}
                      title="Delete client"
                    >
                      <Trash2 style={{ width: 14, height: 14 }} />
                    </button>
                  )}
                  <ChevronRight style={{ width: 16, height: 16, color: '#BC8D26', flexShrink: 0 }} />
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
      )}
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
        style={{ width: '100%', padding: '8px 10px', border: '1px solid rgba(0,40,85,0.25)', background: '#FCF7F2', fontFamily: "'Libre Baskerville', Georgia, serif", fontSize: 13, color: '#002855', cursor: 'pointer', outline: 'none' }}
      >
        {options.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
  );
}

function StatusBadge({ status, channel }) {
  const bg = getStatusColor(status);
  const textColor = (status === 'Contacted' || status === 'Pitched') ? '#002855' : '#FCF7F2';
  return (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 text-[10px] font-display tracking-wider" style={{ background: bg, color: textColor, fontWeight: 600 }}>
      {(status || 'NEW').toUpperCase()}
    </span>
  );
}

// =================== Visits Page ===================
// =================== ORDER HISTORY PAGE ===================
// ── MONTH FILTER COMPONENT ────────────────────────────────────────────────────
// ── B2B CUSTOMS PAGE ─────────────────────────────────────────────────────────
// ── B2B CUSTOMS — Production Tracker ────────────────────────────────────────
// Standalone Kanban-style grid for tracking custom B2B production jobs.
// NOT linked to the clients table — purely a manual progress tracker.
const B2B_CUSTOMS_PRODUCTION_STAGES = ['All goods in warehouse', 'Bottling', 'Complete'];

function YesNoSelect({ value, onChange }) {
  const isYes = value === 'Yes';
  return (
    <select
      value={value || 'No'}
      onChange={e => onChange(e.target.value)}
      style={{
        width: '100%', padding: '7px 8px', border: '1px solid rgba(0,40,85,0.2)',
        background: isYes ? 'rgba(45,134,89,0.12)' : 'rgba(204,35,58,0.07)',
        color: isYes ? '#2d8659' : '#CC233A',
        fontFamily: "'Cinzel',serif", fontSize: 10, fontWeight: 700, letterSpacing: '0.08em',
        cursor: 'pointer', outline: 'none', textAlign: 'center',
      }}>
      <option value="Yes">YES</option>
      <option value="No">NO</option>
    </select>
  );
}

function ProductionStageSelect({ value, onChange }) {
  const stageColor = {
    'All goods in warehouse': '#5A7A99',
    'Bottling': '#BC8D26',
    'Complete': '#2d8659',
  }[value] || '#5A7A99';
  return (
    <select
      value={value || B2B_CUSTOMS_PRODUCTION_STAGES[0]}
      onChange={e => onChange(e.target.value)}
      style={{
        width: '100%', padding: '7px 8px', border: '1px solid rgba(0,40,85,0.2)',
        background: `${stageColor}1A`, color: stageColor,
        fontFamily: "'Cinzel',serif", fontSize: 9, fontWeight: 700, letterSpacing: '0.04em',
        cursor: 'pointer', outline: 'none', textAlign: 'center',
      }}>
      {B2B_CUSTOMS_PRODUCTION_STAGES.map(s => <option key={s} value={s}>{s.toUpperCase()}</option>)}
    </select>
  );
}

function B2BCustomsPage({ rows, onAdd, onUpdate, onDelete, askConfirm }) {
  const [adding, setAdding] = useState(false);
  const [newName, setNewName] = useState('');
  const [savingId, setSavingId] = useState(null);
  const [addError, setAddError] = useState('');
  const [justAddedId, setJustAddedId] = useState(null);
  const nameTimers = useRef({});
  const nameInputRefs = useRef({});

  const handleAdd = async () => {
    if (!newName.trim()) return;
    setAdding(true);
    setAddError('');
    try {
      const created = await onAdd({ customerName: newName.trim() });
      setNewName('');
      // Once the project is created, jump focus straight to its customer
      // name field and highlight the row so the rest of the columns
      // (Deposit Paid, Briefed, Liquid & BSOA, Production, Balance Paid,
      // Ready for Dispatch) are immediately visible and ready to use.
      if (created?.id) {
        setJustAddedId(created.id);
        setTimeout(() => {
          const el = nameInputRefs.current[created.id];
          if (el) {
            el.focus();
            el.select();
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
          }
        }, 60);
        setTimeout(() => setJustAddedId(curr => curr === created.id ? null : curr), 2200);
      }
    } catch (err) {
      console.error('[B2BCustomsPage] add failed', err);
      setAddError(err?.message || 'Could not add project — please try again.');
    }
    setAdding(false);
  };

  // Debounced name typing — update local state instantly, persist after a short pause
  const handleNameChange = (row, value) => {
    onUpdate(row.id, { customerName: value }); // optimistic local update happens inside onUpdate
    if (nameTimers.current[row.id]) clearTimeout(nameTimers.current[row.id]);
    setSavingId(row.id);
    nameTimers.current[row.id] = setTimeout(() => {
      setSavingId(curr => curr === row.id ? null : curr);
    }, 600);
  };

  const columns = [
    { key: 'depositPaid',     label: 'Deposit Paid',          type: 'yesno' },
    { key: 'briefed',         label: 'Amp & Mike & Crazy Briefed', type: 'yesno' },
    { key: 'liquidLinedUp',   label: 'Liquid & BSOA Lined Up', type: 'yesno' },
    { key: 'productionStage', label: 'Production',            type: 'stage' },
    { key: 'balancePaid',     label: 'Balance Paid',          type: 'yesno' },
    { key: 'readyDispatch',   label: 'Ready for Dispatch',    type: 'yesno' },
  ];

  return (
    <div className="fade-up space-y-5">
      {/* Header */}
      <div className="pb-3 border-b">
        <p className="font-display text-[9px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>B2B CHANNEL</p>
        <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', gap: 12, marginTop: 4, flexWrap: 'wrap' }}>
          <div>
            <h1 className="font-display ink" style={{ fontWeight: 700, fontSize: 28 }}>B2B CUSTOMS</h1>
            <p className="italic ocean" style={{ fontSize: 12, marginTop: 2 }}>
              Production tracker — manual progress board, not linked to the client database
            </p>
          </div>
          <span className="font-display text-[9px] tracking-[0.15em] ocean" style={{ fontWeight: 600 }}>
            {rows.length} PROJECT{rows.length !== 1 ? 'S' : ''}
          </span>
        </div>
      </div>

      {/* Grid */}
      <div className="premium-card" style={{ overflowX: 'auto' }}>
        <div style={{ minWidth: 980 }}>
          {/* Header row */}
          <div style={{ display: 'grid', gridTemplateColumns: '220px repeat(6, 1fr) 40px', background: '#002855' }}>
            <div style={{ padding: '12px 14px' }}>
              <p style={{ fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.2em', color: '#DBB85E', fontWeight: 700 }}>CUSTOMER NAME</p>
            </div>
            {columns.map(col => (
              <div key={col.key} style={{ padding: '12px 8px', textAlign: 'center', borderLeft: '1px solid rgba(255,255,255,0.1)' }}>
                <p style={{ fontFamily: "'Cinzel',serif", fontSize: 8, letterSpacing: '0.1em', color: '#FCF7F2', fontWeight: 700, lineHeight: 1.3 }}>{col.label.toUpperCase()}</p>
              </div>
            ))}
            <div></div>
          </div>

          {/* Data rows */}
          {rows.length === 0 ? (
            <div style={{ padding: '32px 16px', textAlign: 'center' }}>
              <p className="italic ocean" style={{ fontSize: 13 }}>No production projects yet — add the first customer below.</p>
            </div>
          ) : (
            rows.map((row, i) => (
              <div key={row.id} style={{
                display: 'grid', gridTemplateColumns: '220px repeat(6, 1fr) 40px',
                borderTop: '1px solid rgba(0,40,85,0.08)',
                background: justAddedId === row.id ? 'rgba(188,141,38,0.14)' : (i % 2 === 0 ? '#FCF7F2' : 'rgba(0,40,85,0.02)'),
                alignItems: 'center',
                transition: 'background 0.8s ease',
              }}>
                {/* Customer name — freely editable */}
                <div style={{ padding: '8px 14px', position: 'relative' }}>
                  <input
                    type="text"
                    ref={el => { if (el) nameInputRefs.current[row.id] = el; }}
                    value={row.customerName}
                    onChange={e => handleNameChange(row, e.target.value)}
                    placeholder="Customer name..."
                    style={{ width: '100%', padding: '6px 8px', border: justAddedId === row.id ? '1px solid #BC8D26' : '1px solid rgba(0,40,85,0.15)', background: '#fff', fontFamily: "'Libre Baskerville',Georgia,serif", fontSize: 13, color: '#002855', fontWeight: 700, outline: 'none' }}
                  />
                  {savingId === row.id && (
                    <span style={{ position: 'absolute', right: 18, top: '50%', transform: 'translateY(-50%)', fontSize: 8, color: '#BC8D26', fontStyle: 'italic' }}>saving…</span>
                  )}
                </div>

                {/* Status columns */}
                {columns.map(col => (
                  <div key={col.key} style={{ padding: '8px' }}>
                    {col.type === 'yesno' ? (
                      <YesNoSelect value={row[col.key]} onChange={(val) => onUpdate(row.id, { [col.key]: val })} />
                    ) : (
                      <ProductionStageSelect value={row[col.key]} onChange={(val) => onUpdate(row.id, { [col.key]: val })} />
                    )}
                  </div>
                ))}

                {/* Delete */}
                <div style={{ display: 'flex', justifyContent: 'center' }}>
                  <button
                    onClick={() => askConfirm({
                      title: 'Remove this project?',
                      message: `${row.customerName || 'Unnamed project'}\n\nThis cannot be undone.`,
                      confirmLabel: 'DELETE',
                      danger: true,
                      onConfirm: () => onDelete(row.id),
                    })}
                    style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'rgba(204,35,58,0.5)', padding: 6, display: 'flex' }}
                    onMouseEnter={e => e.currentTarget.style.color = '#CC233A'}
                    onMouseLeave={e => e.currentTarget.style.color = 'rgba(204,35,58,0.5)'}
                    title="Remove project">
                    <Trash2 style={{ width: 14, height: 14 }} />
                  </button>
                </div>
              </div>
            ))
          )}

          {/* Add new row */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '12px 14px', borderTop: '2px solid rgba(188,141,38,0.3)', background: 'rgba(188,141,38,0.04)', flexWrap: 'wrap' }}>
            <UserPlus style={{ width: 14, height: 14, color: '#BC8D26', flexShrink: 0 }} />
            <input
              type="text"
              value={newName}
              onChange={e => { setNewName(e.target.value); if (addError) setAddError(''); }}
              onKeyDown={e => { if (e.key === 'Enter') handleAdd(); }}
              placeholder="Add new customer / project..."
              style={{ flex: 1, maxWidth: 300, padding: '8px 10px', border: addError ? '1px solid #CC233A' : '1px solid rgba(0,40,85,0.2)', background: '#fff', fontFamily: "'Libre Baskerville',Georgia,serif", fontSize: 13, color: '#002855', outline: 'none' }}
            />
            <button
              onClick={handleAdd}
              disabled={adding || !newName.trim()}
              style={{ padding: '8px 16px', background: adding || !newName.trim() ? 'rgba(0,40,85,0.2)' : '#002855', color: '#FCF7F2', border: 'none', fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.2em', fontWeight: 700, cursor: adding || !newName.trim() ? 'default' : 'pointer' }}>
              {adding ? 'ADDING...' : 'ADD PROJECT'}
            </button>
            {addError && (
              <span style={{ fontSize: 11, color: '#CC233A', fontStyle: 'italic', width: '100%' }}>{addError}</span>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

function MonthFilter({ value, onChange }) {
  // value = 'All' | 'YYYY-MM'
  // Build list: All + last 18 months descending
  const months = useMemo(() => {
    const list = [{ label: 'ALL TIME', value: 'All' }];
    const now = new Date();
    for (let i = 0; i < 18; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const iso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      const label = d.toLocaleString('default', { month: 'short', year: 'numeric' }).toUpperCase();
      list.push({ label, value: iso });
    }
    return list;
  }, []);

  const current = months.find(m => m.value === value) || months[0];

  const prev = () => {
    const idx = months.findIndex(m => m.value === value);
    if (idx < months.length - 1) onChange(months[idx + 1].value);
  };
  const next = () => {
    const idx = months.findIndex(m => m.value === value);
    if (idx > 0) onChange(months[idx - 1].value);
  };

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 0, border: '1px solid rgba(0,40,85,0.2)', overflow: 'hidden' }}>
      <button onClick={prev} style={{ padding: '7px 12px', background: 'none', border: 'none', cursor: 'pointer', fontFamily: "'Cinzel',serif", fontSize: 12, color: '#002855', fontWeight: 700, borderRight: '1px solid rgba(0,40,85,0.15)' }}>‹</button>
      <select
        value={value}
        onChange={e => onChange(e.target.value)}
        style={{ padding: '7px 10px', background: '#FCF7F2', border: 'none', fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.15em', fontWeight: 700, color: '#002855', cursor: 'pointer', outline: 'none', minWidth: 100 }}>
        {months.map(m => <option key={m.value} value={m.value}>{m.label}</option>)}
      </select>
      <button onClick={next} disabled={value === months[1]?.value} style={{ padding: '7px 12px', background: 'none', border: 'none', cursor: value === months[1]?.value ? 'default' : 'pointer', fontFamily: "'Cinzel',serif", fontSize: 12, color: value === months[1]?.value ? 'rgba(0,40,85,0.3)' : '#002855', fontWeight: 700, borderLeft: '1px solid rgba(0,40,85,0.15)' }}>›</button>
    </div>
  );
}

function OrderHistoryPage({ clients, visits, onDeleteVisit }) {
  const [search, setSearch] = useState('');
  const [expandedOrderId, setExpandedOrderId] = useState(null);
  const [filterChannel, setFilterChannel] = useState('All');
  const [filterRep, setFilterRep] = useState('All');
  const [sortBy, setSortBy] = useState('date');
  const [filterMonth, setFilterMonth] = useState('All');

  // Only visits that have actual orders (items with qty > 0)
  const orders = useMemo(() => {
    return visits
      .filter(v => v.items && v.items.length > 0 && v.items.some(it => Number(it.qty) > 0))
      .map(v => {
        const client = clients.find(c => c.id === v.clientId);
        return { ...v, clientName: client?.venue || 'Unknown', clientChannel: client?.channel || '', clientLocation: client?.location || '', clientId: v.clientId };
      })
      .filter(v => {
        if (filterChannel !== 'All' && v.clientChannel !== filterChannel) return false;
        if (filterRep !== 'All' && v.salesRep !== filterRep) return false;
        if (filterMonth !== 'All' && !(v.date || '').startsWith(filterMonth)) return false;
        if (search.trim()) {
          const q = search.trim().toLowerCase();
          return v.clientName.toLowerCase().includes(q) ||
            (v.salesRep || '').toLowerCase().includes(q) ||
            (v.date || '').includes(q) ||
            v.items.some(it => (it.name || '').toLowerCase().includes(q));
        }
        return true;
      })
      .sort((a, b) => sortBy === 'amount'
        ? (b.saleAmount || 0) - (a.saleAmount || 0)
        : (b.date || '').localeCompare(a.date || ''));
  }, [visits, clients, search, filterChannel, filterRep, filterMonth, sortBy]);

  const totalRevenue = orders.reduce((s, o) => s + (o.saleAmount || 0), 0);
  const statusColors = (s) => getStatusColor(s);

  return (
    <div className="fade-up space-y-5">
      {/* Header */}
      <div className="pb-3 border-b">
        <p className="font-display text-[9px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>ORDER DATABASE</p>
        <div style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'space-between', gap: 12, marginTop: 4, flexWrap: 'wrap' }}>
          <div>
            <h1 className="font-display ink" style={{ fontWeight: 700, fontSize: 28 }}>CLIENT ORDER HISTORY</h1>
            <p className="italic ocean" style={{ fontSize: 12, marginTop: 2 }}>
              {orders.length} orders · {ZAR(totalRevenue)} total revenue
            </p>
          </div>
        </div>
      </div>

      {/* Filters */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, border: '1px solid rgba(0,40,85,0.25)', background: '#FCF7F2', padding: '0 12px' }}>
          <Search style={{ width: 16, height: 16, color: '#5A7A99', flexShrink: 0 }} />
          <input
            type="text"
            placeholder="Search venue, rep, SKU, date..."
            value={search}
            onChange={e => setSearch(e.target.value)}
            style={{ flex: 1, border: 'none', background: 'transparent', padding: '12px 0', fontFamily: "'Libre Baskerville', Georgia, serif", fontSize: 14, color: '#002855', outline: 'none' }}
          />
          {search && <button onClick={() => setSearch('')} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#5A7A99', fontSize: 18, lineHeight: 1 }}>×</button>}
        </div>
        {/* Channel filter */}
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          {['All', ...CHANNELS].map(ch => (
            <button key={ch} onClick={() => setFilterChannel(ch)}
              style={{ padding: '6px 14px', border: `1px solid ${filterChannel === ch ? '#002855' : 'rgba(0,40,85,0.2)'}`, background: filterChannel === ch ? '#002855' : 'none', color: filterChannel === ch ? '#FCF7F2' : '#002855', fontFamily: "'Cinzel', serif", fontSize: 9, letterSpacing: '0.15em', fontWeight: 600, cursor: 'pointer' }}>
              {ch === 'All' ? 'ALL' : ch === 'B2B' ? 'PRIVATE' : ch === 'Trade' ? 'RETAIL' : ch.toUpperCase()}
            </button>
          ))}
        </div>
        {/* Month + Rep + Sort row */}
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
          <span style={{ fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.2em', color: '#5A7A99', fontWeight: 600 }}>MONTH:</span>
          <MonthFilter value={filterMonth} onChange={setFilterMonth} />
          <span style={{ fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.2em', color: '#5A7A99', fontWeight: 600, marginLeft: 8 }}>AGENT:</span>
          {['All', ...SALES_REPS].map(r => (
            <button key={r} onClick={() => setFilterRep(r)}
              style={{ padding: '5px 12px', border: `1px solid ${filterRep === r ? '#BC8D26' : 'rgba(0,40,85,0.2)'}`, background: filterRep === r ? '#BC8D26' : 'none', color: filterRep === r ? '#FCF7F2' : '#002855', fontFamily: "'Cinzel', serif", fontSize: 9, letterSpacing: '0.12em', fontWeight: 600, cursor: 'pointer' }}>
              {r === 'All' ? 'ALL' : r.toUpperCase()}
            </button>
          ))}
          <div style={{ marginLeft: 'auto', display: 'flex', gap: 6 }}>
            <span style={{ fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.2em', color: '#5A7A99', fontWeight: 600, alignSelf: 'center' }}>SORT:</span>
            {[{ key: 'date', label: 'DATE' }, { key: 'amount', label: 'AMOUNT' }].map(s => (
              <button key={s.key} onClick={() => setSortBy(s.key)}
                style={{ padding: '5px 12px', border: `1px solid ${sortBy === s.key ? '#002855' : 'rgba(0,40,85,0.2)'}`, background: sortBy === s.key ? '#002855' : 'none', color: sortBy === s.key ? '#FCF7F2' : '#002855', fontFamily: "'Cinzel', serif", fontSize: 9, letterSpacing: '0.12em', fontWeight: 600, cursor: 'pointer' }}>
                {s.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Order list */}
      {orders.length === 0 ? (
        <div className="premium-card" style={{ padding: 40, textAlign: 'center' }}>
          <p style={{ fontStyle: 'italic', color: '#5A7A99', fontSize: 14 }}>No orders found{search ? ` matching "${search}"` : ''}.</p>
        </div>
      ) : (
        <div className="premium-card" style={{ overflow: 'hidden' }}>
          {orders.map((order, i) => {
            const isOpen = expandedOrderId === order.id;
            const skuCount = order.items.length;
            const totalQty = order.items.reduce((s, it) => s + (Number(it.qty) || 0), 0);
            return (
              <div key={order.id} style={{ borderTop: i === 0 ? 'none' : '1px solid rgba(0,40,85,0.1)' }}>

                {/* Order summary row */}
                <div style={{ display: 'flex', alignItems: 'center', background: isOpen ? 'rgba(0,40,85,0.04)' : '#FCF7F2', borderLeft: `3px solid ${order.clientChannel === 'B2B' ? '#BC8D26' : '#5A7A99'}`, transition: 'background 0.12s' }}>
                  <div
                    onClick={() => setExpandedOrderId(isOpen ? null : order.id)}
                    style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '13px 12px 13px 16px', cursor: 'pointer', flex: 1, minWidth: 0 }}
                  >
                    {/* Date */}
                    <div style={{ flexShrink: 0, minWidth: 80 }}>
                      <p className="font-display ink" style={{ fontWeight: 700, fontSize: 12 }}>{order.date || '—'}</p>
                      <p style={{ fontSize: 9, color: '#5A7A99', fontStyle: 'italic' }}>{order.salesRep}</p>
                    </div>
                    {/* Venue */}
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <p className="font-display ink" style={{ fontWeight: 700, fontSize: 13, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{order.clientName}</p>
                      <p style={{ fontSize: 10, color: '#5A7A99', fontStyle: 'italic' }}>
                        {order.clientLocation || ''}{order.clientLocation && order.clientChannel ? ' · ' : ''}{order.clientChannel === 'B2B' ? 'Private' : order.clientChannel === 'Trade' ? 'Retail' : ''}
                      </p>
                    </div>
                    {/* SKU summary */}
                    <div style={{ flexShrink: 0, textAlign: 'right' }}>
                      <p style={{ fontSize: 10, color: '#5A7A99', fontStyle: 'italic' }}>{skuCount} SKU{skuCount !== 1 ? 's' : ''} · {totalQty} units</p>
                    </div>
                    {/* Total */}
                    <div style={{ flexShrink: 0, minWidth: 72, textAlign: 'right' }}>
                      <p className="font-display copper" style={{ fontWeight: 700, fontSize: 13 }}>{ZAR(order.saleAmount || 0)}</p>
                    </div>
                    <ChevronDown style={{ width: 16, height: 16, color: '#BC8D26', flexShrink: 0, transform: isOpen ? 'rotate(180deg)' : 'none', transition: 'transform 0.2s' }} />
                  </div>
                  {/* Delete button */}
                  <button
                    type="button"
                    onClick={() => { onDeleteVisit && onDeleteVisit(order.id); }}
                    style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '13px 14px', color: 'rgba(0,40,85,0.2)', flexShrink: 0, transition: 'color 0.15s' }}
                    onMouseEnter={e => e.currentTarget.style.color = '#CC233A'}
                    onMouseLeave={e => e.currentTarget.style.color = 'rgba(0,40,85,0.2)'}
                    title="Delete order"
                  >
                    <Trash2 style={{ width: 14, height: 14 }} />
                  </button>
                </div>

                {/* Expanded SKU breakdown */}
                {isOpen && (
                  <div style={{ background: 'rgba(0,40,85,0.02)', borderTop: '1px solid rgba(0,40,85,0.08)', padding: '0 16px 16px' }}>
                    {/* Order meta */}
                    <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap', padding: '12px 0 10px', borderBottom: '1px solid rgba(0,40,85,0.08)', marginBottom: 10 }}>
                      <div>
                        <p style={{ fontSize: 8, letterSpacing: '0.2em', color: '#BC8D26', fontFamily: "'Cinzel',serif", fontWeight: 600 }}>DATE</p>
                        <p className="font-display ink" style={{ fontSize: 12, fontWeight: 700 }}>{order.date || '—'}</p>
                      </div>
                      <div>
                        <p style={{ fontSize: 8, letterSpacing: '0.2em', color: '#BC8D26', fontFamily: "'Cinzel',serif", fontWeight: 600 }}>SALES REP</p>
                        <p className="font-display ink" style={{ fontSize: 12, fontWeight: 700 }}>{order.salesRep || '—'}</p>
                      </div>
                      <div>
                        <p style={{ fontSize: 8, letterSpacing: '0.2em', color: '#BC8D26', fontFamily: "'Cinzel',serif", fontWeight: 600 }}>OUTCOME</p>
                        <p className="font-display ink" style={{ fontSize: 12, fontWeight: 700 }}>{order.outcome || '—'}</p>
                      </div>
                      {order.notes && (
                        <div style={{ flex: 1 }}>
                          <p style={{ fontSize: 8, letterSpacing: '0.2em', color: '#BC8D26', fontFamily: "'Cinzel',serif", fontWeight: 600 }}>NOTES</p>
                          <p style={{ fontSize: 11, color: '#002855', fontStyle: 'italic' }}>{order.notes}</p>
                        </div>
                      )}
                    </div>

                    {/* SKU table */}
                    <div>
                      {/* Header */}
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 60px 80px 80px 80px', gap: 8, padding: '6px 8px', background: '#002855' }}>
                        {['SKU', 'QTY', 'UNIT R', 'DISC', 'LINE'].map(h => (
                          <p key={h} style={{ fontSize: 8, letterSpacing: '0.2em', color: '#DBB85E', fontFamily: "'Cinzel',serif", fontWeight: 600, textAlign: h !== 'SKU' ? 'right' : 'left' }}>{h}</p>
                        ))}
                      </div>
                      {/* Rows */}
                      {order.items.map((it, idx) => {
                        const lineTotal = (Number(it.unitPrice) || 0) * (Number(it.qty) || 0);
                        const discounted = Number(it.listPrice) > Number(it.unitPrice);
                        const discAmt = discounted ? (Number(it.listPrice) - Number(it.unitPrice)) * Number(it.qty) : 0;
                        return (
                          <div key={idx} style={{ display: 'grid', gridTemplateColumns: '1fr 60px 80px 80px 80px', gap: 8, padding: '9px 8px', borderBottom: idx < order.items.length - 1 ? '1px solid rgba(0,40,85,0.08)' : 'none', background: idx % 2 === 0 ? '#FCF7F2' : 'rgba(0,40,85,0.02)' }}>
                            <div>
                              <p className="font-display ink" style={{ fontWeight: 700, fontSize: 12 }}>{it.name}</p>
                              {discounted && <p style={{ fontSize: 9, color: '#BC8D26', fontStyle: 'italic' }}>from {ZAR(it.listPrice)}</p>}
                            </div>
                            <p style={{ fontSize: 12, fontWeight: 700, color: '#002855', textAlign: 'right', fontFamily: "'Cinzel',serif" }}>{it.qty}</p>
                            <p style={{ fontSize: 12, fontWeight: discounted ? 700 : 400, color: discounted ? '#BC8D26' : '#002855', textAlign: 'right', fontFamily: "'Cinzel',serif" }}>{ZAR(it.unitPrice)}</p>
                            <p style={{ fontSize: 12, color: discounted ? '#BC8D26' : 'rgba(0,40,85,0.3)', textAlign: 'right', fontFamily: "'Cinzel',serif" }}>{discounted ? `-${ZAR(discAmt)}` : '—'}</p>
                            <p style={{ fontSize: 12, fontWeight: 700, color: '#002855', textAlign: 'right', fontFamily: "'Cinzel',serif" }}>{ZAR(lineTotal)}</p>
                          </div>
                        );
                      })}
                      {/* Total row */}
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 60px 80px 80px 80px', gap: 8, padding: '10px 8px', background: '#002855', marginTop: 1 }}>
                        <p style={{ fontSize: 9, letterSpacing: '0.2em', color: '#DBB85E', fontFamily: "'Cinzel',serif", fontWeight: 700, gridColumn: '1 / 5', textAlign: 'right' }}>TOTAL EX VAT</p>
                        <p style={{ fontSize: 13, fontWeight: 700, color: '#DBB85E', textAlign: 'right', fontFamily: "'Cinzel',serif" }}>{ZAR(order.saleAmount || 0)}</p>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function VisitsPage({ visits, clients, onLog, onEdit, onDelete, onEmail }) {
  const [filterRep, setFilterRep] = useState('All');
  const [filterMonth, setFilterMonth] = useState('All');
  const filtered = visits
    .filter(v => filterRep === 'All' || v.salesRep === filterRep)
    .filter(v => filterMonth === 'All' || (v.date || '').startsWith(filterMonth))
    .sort((a, b) => (b.date || '').localeCompare(a.date || ''));

  return (
    <div className="space-y-6 fade-up">
      {/* Page header */}
      <div className="pb-3 border-b">
        <p className="font-display text-[9px] tracking-[0.4em] copper" style={{ fontWeight: 600 }}>VISIT HISTORY</p>
        <div className="flex items-center justify-between gap-3 mt-1 flex-wrap">
          <h1 className="font-display text-2xl md:text-3xl ink" style={{ fontWeight: 700 }}>VISIT LOG</h1>
          <div className="flex items-center gap-3 flex-wrap">
            <MonthFilter value={filterMonth} onChange={setFilterMonth} />
            <RepToggle active={filterRep} onChange={setFilterRep} />
            <button onClick={onLog}
              className="flex items-center gap-2 px-4 py-2.5 font-display text-[10px] tracking-[0.2em]"
              style={{ background: '#BC8D26', color: '#FCF7F2', fontWeight: 700 }}>
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
          <button onClick={onLog} className="bg-copper px-6 py-3 font-display text-xs tracking-[0.25em]" style={{ color: '#FCF7F2', fontWeight: 700 }}>LOG YOUR FIRST VISIT</button>
        </div>
      ) : (
        <div className="space-y-2">
          {filtered.map(v => {
            const c = clients.find(c => c.id === v.clientId);
            const outcomeColor = v.outcome === 'Sold In' ? '#2d8659' : v.outcome === 'Rejected' ? '#CC233A' : '#BC8D26';
            return (
              <div key={v.id} className="premium-card border-l-4" style={{ borderLeftColor: outcomeColor }}>
                {/* Main row — always visible */}
                <div className="p-3 md:p-5 flex items-start justify-between gap-3">
                  <div className="flex-1 min-w-0">
                    {/* Venue + outcome */}
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="font-display ink text-sm truncate" style={{ fontWeight: 700 }}>{c?.venue || 'Unknown'}</span>
                      <span className="text-[9px] font-display tracking-wider px-1.5 py-0.5 flex-shrink-0" style={{ background: outcomeColor, color: '#FCF7F2', fontWeight: 600 }}>{v.outcome.toUpperCase()}</span>
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
                    <button type="button" onClick={() => onEmail(v)} className="p-2 hover: border border" title="Email order">
                      <Mail className="w-3.5 h-3.5" style={{ color: '#5A7A99' }} />
                    </button>
                    <button type="button" onClick={() => onEdit(v)} className="p-2 hover: border border">
                      <Edit2 className="w-3.5 h-3.5 ocean" />
                    </button>
                    <button type="button" onClick={() => onDelete(v.id)} className="p-2  border border">
                      <Trash2 className="w-3.5 h-3.5" style={{ color: '#CC233A' }} />
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
function LogVisitModal({ clients, onClose, onSubmit, onRequestNewClient, existingVisit, skuOverrides = {}, preselectedClientId = null, preselectedSalesRep = '' }) {
  const effectiveSkus = SKU_CATALOGUE.map(s => ({ ...s, price: skuOverrides[s.id] ?? s.price }));
  const isEdit = !!existingVisit;
  const isEmailOnly = existingVisit?._emailOnly === true;
  const [salesRep, setSalesRep] = useState(existingVisit?.salesRep || preselectedSalesRep || 'Alex');
  const [clientId, setClientId] = useState(existingVisit?.clientId ? String(existingVisit.clientId) : preselectedClientId ? String(preselectedClientId) : '');
  const [date, setDate] = useState(existingVisit?.date || todayISO());
  const [outcome, setOutcome] = useState(existingVisit?.outcome || 'Met / Discussion');
  const [contactMethod, setContactMethod] = useState(existingVisit?.contactMethod || 'In Person');
  const [followUpDate, setFollowUpDate] = useState(existingVisit?.followUpDate || '');
  const [items, setItems] = useState(existingVisit?.items ? existingVisit.items.map(it => ({ ...it })) : []);
  const [notes, setNotes] = useState(existingVisit?.notes || '');
  const [followUp, setFollowUp] = useState(existingVisit?.followUp || '');
  const [taggedReps, setTaggedReps] = useState(existingVisit?.taggedReps || []);
  // Auto-open email modal when opened in emailOnly mode
  const [emailModalOpen, setEmailModalOpen] = useState(isEmailOnly);
  const [search, setSearch] = useState('');
  const [skuPickerOpen, setSkuPickerOpen] = useState(false);
  const [showCloseConfirm, setShowCloseConfirm] = useState(false);

  // Has the user entered any data worth warning about?
  const logIsDirty = !isEdit && (!!clientId || !!notes.trim() || items.length > 0);

  const handleLogClose = () => {
    if (logIsDirty) { setShowCloseConfirm(true); } else { onClose(); }
  };

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
    lines.push('Dear,');
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
    lines.push(`Contact Method: ${contactMethod || 'In Person'}`);
    lines.push(`Outcome:        ${outcome}`);
    lines.push(`Sales Rep:      ${salesRep} — Avante Cape Brandy`);
    if (c?.email && !isBlockedEmail(c.email)) lines.push(`Email Invoice:  ${c.email}`);
    const terms = c?.paymentTerms || 'COD';
    lines.push(`Payment Terms:  ${terms === 'COD' ? 'Cash on Delivery (COD)' : terms === '30 Days' ? '30 Days from Invoice' : terms === '60 Days' ? '60 Days from Invoice' : terms}`);
    if (followUpDate) lines.push(`Follow-up Due:  ${followUpDate}`);
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

    // Build recipient list per channel: Private Sales & B2B → matthew,
    // Trade Retail / other channels → orders@redbev. Client email always included if on file.
    const recipientParts = [];
    if (c?.email && !isBlockedEmail(c.email)) recipientParts.push(c.email);
    const isMatthewChannel = c?.channel === 'B2B';
    const defaultAddr = isMatthewChannel ? 'matthew@breakfreebeverages.com' : 'orders@redbev.co.za';
    if (!recipientParts.includes(defaultAddr)) recipientParts.push(defaultAddr);
    const defaultRecipient = recipientParts.join(', ');

    return { subject, body: lines.join('\n'), defaultRecipient, orderTotal };
  };

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

  const [submitting, setSubmitting] = useState(false);

  const handleSave = async () => {
    if (submitting) return; // prevent double-tap
    setValidationError('');
    if (!clientId) {
      setValidationError('Please select a client before saving.');
      return;
    }
    setSubmitting(true);
    const payload = {
      salesRep,
      clientId: Number(clientId),
      channel: selectedClient?.channel || '',
      date, outcome, contactMethod, followUpDate,
      items: items.map(it => ({
        skuId: it.skuId,
        name: it.name,
        qty: Number(it.qty) || 0,
        unitPrice: Number(it.unitPrice) || 0,
        listPrice: Number(it.listPrice) || 0,
        lineTotal: (Number(it.unitPrice) || 0) * (Number(it.qty) || 0),
      })),
      notes, followUp, taggedReps,
    };
    try {
      notifyNewTags({
        prevTags: existingVisit?.taggedReps || [],
        newTags: taggedReps || [],
        clientName: selectedClient?.venue || '',
        location: selectedClient?.location || '',
        channel: selectedClient?.channel || '',
        taggedBy: salesRep,
        type: 'visit',
        date, outcome, notes, followUp,
      });
      await onSubmit(payload);
    } catch (err) {
      console.error('[LogVisit] onSubmit threw:', err);
      setValidationError('Save failed: ' + (err?.message || String(err)));
      setSubmitting(false); // re-enable on error so they can retry
    }
    // Don't re-enable on success — modal closes
  };

  const requestAddClient = () => {
    onRequestNewClient(salesRep, (created) => {
      setClientId(String(created.id));
    });
  };

  return (
    <div onClick={handleLogClose} style={{ position:"fixed", inset:0, zIndex:50, display:"flex", alignItems:"flex-start", justifyContent:"center", padding:'8px', background:"rgba(0,40,85,0.78)", backdropFilter:"blur(4px)", overflowY:"auto", overflowX:"hidden" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background:"#FCF7F2", width:"100%", maxWidth:560, maxHeight:"94vh", overflowY:"auto", overflowX:"hidden", border:"2px solid #BC8D26", position:"relative", margin:"8px 0", flexShrink:0 }}>
        <div className="absolute top-0 left-0 right-0 h-1 bg-copper z-10"></div>

        {/* Unsaved warning */}
        {showCloseConfirm && (
          <div style={{ position:'fixed', inset:0, zIndex:200, display:'flex', alignItems:'center', justifyContent:'center', background:'rgba(0,40,85,0.85)', padding:20 }} onClick={e => e.stopPropagation()}>
            <div className="premium-card" style={{ maxWidth:320, width:'100%', padding:28, textAlign:'center' }}>
              <p className="font-display" style={{ fontSize:11, letterSpacing:'0.3em', color:'#BC8D26', fontWeight:600, marginBottom:8 }}>DISCARD VISIT?</p>
              <p style={{ fontSize:13, color:'#002855', marginBottom:20 }}>You have unsaved data in this visit log. Are you sure you want to close?</p>
              <div style={{ display:'flex', gap:8 }}>
                <button onClick={() => setShowCloseConfirm(false)}
                  style={{ flex:1, padding:'10px', border:'1px solid rgba(0,40,85,0.2)', background:'none', fontFamily:"'Cinzel',serif", fontSize:9, letterSpacing:'0.2em', color:'#002855', cursor:'pointer', fontWeight:600 }}>
                  KEEP EDITING
                </button>
                <button onClick={() => { setShowCloseConfirm(false); onClose(); }}
                  style={{ flex:1, padding:'10px', background:'#CC233A', border:'none', fontFamily:"'Cinzel',serif", fontSize:9, letterSpacing:'0.2em', color:'#FCF7F2', cursor:'pointer', fontWeight:700 }}>
                  DISCARD
                </button>
              </div>
            </div>
          </div>
        )}
        {/* Header */}
        <div className="p-4 md:p-6 relative overflow-hidden" style={{ background: '#002855' }}>
          <RaysBackdrop opacity={0.06} />
          <div className="flex items-center justify-between gap-4 relative">
            <div className="flex items-center gap-3">
              <AvanteLogo height={36} />
              <div>
                <p className="font-display text-[9px] tracking-[0.3em]" style={{ color: '#DBB85E', fontWeight: 600 }}>{isEdit ? 'AMEND RECORD' : 'DARE TO FORWARD'}</p>
                <h2 className="font-display text-xl md:text-2xl mt-0.5" style={{ color: '#FCF7F2', fontWeight: 700, letterSpacing: '0.08em' }}>{isEdit ? 'EDIT VISIT' : 'LOG A VISIT'}</h2>
              </div>
            </div>
            <button type="button" onClick={handleLogClose} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'rgba(255,255,255,0.65)', padding: '4px 8px', fontSize: 22, fontWeight: 200, lineHeight: 1, flexShrink: 0 }}>✕</button>
          </div>
        </div>

        <div className="p-4 md:p-6 space-y-4">
          {/* Sales rep */}
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>SALES MEMBER</label>
            <div className="grid grid-cols-3 md:grid-cols-5 gap-2">
              {SALES_REPS.map(r => (
                <button key={r} type="button" onClick={() => { setSalesRep(r); setClientId(''); }} className={`py-3 font-display text-xs tracking-[0.2em] border transition-all ${salesRep === r ? 'bg-ink border-ink' : 'bg-cream border hover:border-copper'}`} style={{ color: salesRep === r ? '#FCF7F2' : '#002855', fontWeight: 700 }}>
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

          {/* Contact Method + Follow-up Date */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>CONTACT METHOD</label>
              <select value={contactMethod} onChange={(e) => setContactMethod(e.target.value)} className="w-full px-3 py-3 border border bg-cream font-body text-sm focus:outline-none focus:border-copper">
                {CONTACT_METHODS.map(m => <option key={m}>{m}</option>)}
              </select>
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>FOLLOW-UP DATE</label>
              <input type="date" value={followUpDate} onChange={(e) => setFollowUpDate(e.target.value)} className="w-full px-3 py-3 border border bg-cream font-body text-sm focus:outline-none focus:border-copper" style={{ color: followUpDate ? '#002855' : 'rgba(0,40,85,0.4)' }} />
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
                    const sku = effectiveSkus.find(s => s.id === e.target.value);
                    if (sku) addSku(sku);
                    e.target.value = '';
                  }}
                  style={{ padding: '6px 10px', border: '1px solid rgba(0,40,85,0.25)', background: '#FCF7F2', fontFamily: "'Cinzel', serif", fontSize: 10, letterSpacing: '0.12em', color: '#002855', cursor: 'pointer', fontWeight: 700 }}
                >
                  <option value="">+ ADD SKU</option>
                  {effectiveSkus.map(sku => (
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
                    const discountPct = it.listPrice > 0 ? Math.round((1 - Number(it.unitPrice) / Number(it.listPrice)) * 100) : 0;
                    const discounted = Number(it.unitPrice) < Number(it.listPrice);
                    return (
                      <div key={it.skuId} className="p-3">
                        <div className="flex items-start justify-between gap-2 mb-2">
                          <div>
                            <p className="ink text-sm font-display" style={{ fontWeight: 700 }}>{it.name}</p>
                            {discounted && <p className="text-[10px] copper italic">disc. from {ZAR(it.listPrice)} · saving {ZAR(it.listPrice - it.unitPrice)}/unit</p>}
                          </div>
                          <button type="button" onClick={() => removeItem(it.skuId)} className="p-1 text-ink/40 hover:text-red-700 flex-shrink-0">
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                        <div className="grid grid-cols-4 gap-2">
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
                              style={{ color: discounted ? '#BC8D26' : '#002855' }} />
                          </div>
                          <div>
                            <p className="font-display text-[9px] tracking-[0.15em] ocean mb-1" style={{ fontWeight: 600 }}>DISC %</p>
                            <input type="number" min="0" max="100" step="1"
                              value={discountPct === 0 ? '' : discountPct}
                              placeholder="0"
                              onChange={(e) => {
                                const pct = Math.max(0, Math.min(100, Number(e.target.value) || 0));
                                const newPrice = Number(it.listPrice) * (1 - pct / 100);
                                updateItem(it.skuId, 'unitPrice', newPrice.toFixed(2));
                              }}
                              className="w-full px-2 py-2 border border bg-cream text-sm text-center focus:outline-none focus:border-copper"
                              style={{ color: discountPct > 0 ? '#BC8D26' : '#002855' }} />
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
                  <div className="grid grid-cols-12 gap-2 px-3 py-2  border-b">
                    <div className="col-span-4 font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>SKU</div>
                    <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-center" style={{ fontWeight: 600 }}>QTY</div>
                    <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-right" style={{ fontWeight: 600 }}>UNIT R</div>
                    <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-center" style={{ fontWeight: 600 }}>DISC %</div>
                    <div className="col-span-1 font-display text-[10px] tracking-[0.2em] ocean text-right" style={{ fontWeight: 600 }}>LINE</div>
                    <div className="col-span-1"></div>
                  </div>
                  {items.map(it => {
                    const lineTotal = (Number(it.unitPrice) || 0) * (Number(it.qty) || 0);
                    const discountPct = it.listPrice > 0 ? Math.round((1 - Number(it.unitPrice) / Number(it.listPrice)) * 100) : 0;
                    const discounted = Number(it.unitPrice) < Number(it.listPrice);
                    return (
                      <div key={it.skuId} className="grid grid-cols-12 gap-2 px-3 py-2 border-b items-center">
                        <div className="col-span-4">
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
                            style={{ color: discounted ? '#BC8D26' : '#002855' }} />
                        </div>
                        <div className="col-span-2">
                          <input type="number" min="0" max="100" step="1"
                            value={discountPct === 0 ? '' : discountPct}
                            placeholder="0%"
                            onChange={(e) => {
                              const pct = Math.max(0, Math.min(100, Number(e.target.value) || 0));
                              const newPrice = Number(it.listPrice) * (1 - pct / 100);
                              updateItem(it.skuId, 'unitPrice', newPrice.toFixed(2));
                            }}
                            className="w-full px-2 py-1 border border bg-cream text-xs text-center focus:outline-none focus:border-copper"
                            style={{ color: discountPct > 0 ? '#BC8D26' : '#002855' }} />
                        </div>
                        <div className="col-span-1 text-right">
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
                <div className="grid grid-cols-12 gap-2 px-3 py-3" style={{ background: '#002855' }}>
                  <div className="col-span-8 font-display text-[10px] tracking-[0.3em]" style={{ color: '#DBB85E', fontWeight: 600 }}>ORDER TOTAL (EX VAT)</div>
                  <div className="col-span-4 text-right font-display text-base" style={{ color: '#FCF7F2', fontWeight: 700 }}>{ZAR(orderTotal)}</div>
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

          {/* Tag agents */}
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>TAG AGENTS</label>
            <p style={{ fontSize: 10, color: '#5A7A99', fontStyle: 'italic', marginBottom: 8 }}>Tag team members who should be notified or are involved in this visit.</p>
            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
              {SALES_REPS.map(rep => {
                const tagged = (taggedReps || []).includes(rep);
                return (
                  <button key={rep} type="button"
                    onClick={() => setTaggedReps(prev =>
                      tagged ? prev.filter(r => r !== rep) : [...(prev || []), rep]
                    )}
                    style={{ padding: '7px 14px', fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.2em', fontWeight: 700, border: '1px solid', borderColor: tagged ? '#BC8D26' : 'rgba(0,40,85,0.2)', background: tagged ? '#BC8D26' : 'transparent', color: tagged ? '#FCF7F2' : '#002855', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6 }}>
                    {tagged && <span style={{ fontSize: 10 }}>✓</span>} @{rep.toUpperCase()}
                  </button>
                );
              })}
            </div>
            {(taggedReps || []).length > 0 && (
              <p style={{ fontSize: 10, color: '#BC8D26', fontStyle: 'italic', marginTop: 6 }}>
                Tagged: {taggedReps.map(r => `@${r}`).join(', ')}
              </p>
            )}
          </div>

          {/* Actions */}
          {validationError && (
            <div className="px-3 py-2 border-l-4 text-xs" style={{ borderLeftColor: '#CC233A', background: 'rgba(204,35,58,0.08)', color: '#CC233A' }}>
              <span className="font-display tracking-wider" style={{ fontWeight: 700 }}>ERROR:</span> {validationError}
            </div>
          )}
          {/* Mobile: stacked full-width buttons. Desktop: row layout */}
          <div className="pt-4 border-t space-y-2 md:space-y-0 md:flex md:items-center md:justify-between md:gap-3 md:flex-wrap">
            <button type="button" onClick={handleEmailOrder}
              className="w-full md:w-auto flex items-center justify-center gap-2 px-4 py-3 font-display text-xs tracking-[0.25em] border"
              style={{ color: '#5A7A99', fontWeight: 700, borderColor: '#5A7A99' }}>
              <Mail className="w-4 h-4" /> EMAIL ORDER
            </button>
            <div className="flex items-center gap-2">
              <button type="button" onClick={onClose} className="flex-1 md:flex-none px-5 py-3 font-display text-xs tracking-[0.25em] ink border border" style={{ fontWeight: 700 }}>CANCEL</button>
              <button type="button" onClick={handleSave} disabled={submitting}
                className="flex-1 md:flex-none bg-ink px-6 py-3 font-display text-xs tracking-[0.25em] flex items-center justify-center gap-2"
                style={{ color: '#FCF7F2', fontWeight: 700, opacity: submitting ? 0.6 : 1, cursor: submitting ? 'not-allowed' : 'pointer' }}>
                {submitting ? 'SAVING...' : (isEdit ? 'UPDATE' : 'SAVE')} {!submitting && <ArrowUpRight className="w-4 h-4" />}
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
    // Route by channel: Private Sales & B2B → matthew, Trade Retail (& others) → orders@redbev
    const isMatthewChannel = client?.channel === 'B2B';
    const list = isMatthewChannel
      ? ['matthew@breakfreebeverages.com']
      : ['orders@redbev.co.za'];
    // Always add the client's own email if on file
    if (client?.email && !isBlockedEmail(client.email) && !list.includes(client.email)) list.push(client.email);
    return list.join(', ');
  }, [client]);

  const [recipient, setRecipient] = useState(initialRecipients);
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
    <div style={{ position:"fixed", inset:0, zIndex:70, display:"flex", alignItems:"center", justifyContent:"center", padding:16, background:"rgba(0,40,85,0.82)", backdropFilter:"blur(4px)", overflowY:"auto" }} onClick={onClose}>
      <div style={{ background:"#FCF7F2", width:"100%", maxWidth:520, margin:"16px 0", border:"2px solid #BC8D26" }} onClick={(e) => e.stopPropagation()}>
        <div className="p-5 flex items-center justify-between gap-4" style={{ background: '#002855' }}>
          <div className="flex items-center gap-3">
            <Mail className="w-5 h-5" style={{ color: '#DBB85E' }} />
            <div>
              <p className="font-display text-[10px] tracking-[0.4em]" style={{ color: '#DBB85E', fontWeight: 600 }}>SEND ORDER CONFIRMATION</p>
              <h2 className="font-display text-xl mt-0.5" style={{ color: '#FCF7F2', fontWeight: 700, letterSpacing: '0.06em' }}>EMAIL ORDER</h2>
            </div>
          </div>
          <button type="button" onClick={onClose} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'rgba(255,255,255,0.65)', padding: '4px 8px', fontSize: 22, fontWeight: 200, lineHeight: 1, flexShrink: 0 }}>✕</button>
        </div>

        <div className="p-6 space-y-5">
          {/* From */}
          <div>
            <p className="font-display text-[10px] tracking-[0.3em] copper mb-1" style={{ fontWeight: 600 }}>FROM</p>
            <div className="flex items-center gap-3 p-3  border border">
              <div className="w-8 h-8 flex items-center justify-center bg-ink">
                <span className="font-display text-sm" style={{ color: '#DBB85E', fontWeight: 700 }}>{salesRep[0]}</span>
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
                {client?.email && !isBlockedEmail(client.email) && (
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
            <p className="text-[10px] italic ocean mt-1">Private Sales & B2B → matthew@breakfreebeverages.com · Trade Retail / other channels → orders@redbev.co.za</p>
            {client?.venue && (
              <p className="text-[11px] italic ocean mt-1.5">
                Order for <strong className="ink">{client.venue}</strong>
                {client.firstName || client.lastName ? ` · attn. ${[client.firstName, client.lastName].filter(Boolean).join(' ')}` : ''}
              </p>
            )}
            {error && (
              <div className="mt-2 px-3 py-2 border-l-4 text-xs" style={{ borderLeftColor: '#CC233A', background: 'rgba(204,35,58,0.08)', color: '#CC233A' }}>
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

          <div>
            <p className="font-display text-[10px] tracking-[0.3em] copper mb-2" style={{ fontWeight: 600 }}>SEND VIA</p>
            <a
              href={(() => {
                const list = parseRecipients(recipient).filter(e => isValidEmail(e));
                if (list.length === 0) return '#';
                const to = encodeURIComponent(list.join(','));
                const su = encodeURIComponent(composed.subject);
                const bo = encodeURIComponent(composed.body);
                // googlegmail:// is the scheme registered by the Gmail app on iOS
                return `googlegmail://co?to=${to}&subject=${su}&body=${bo}`;
              })()}
              onClick={(e) => {
                const list = parseRecipients(recipient);
                if (list.length === 0) { e.preventDefault(); setError('Please enter at least one recipient.'); return; }
                const invalid = list.filter(em => !isValidEmail(em));
                if (invalid.length > 0) { e.preventDefault(); setError(`Invalid: ${invalid.join(', ')}`); return; }
                setError('');

                const to = list.join(',');
                const su = composed.subject;
                const bo = composed.body;

                // Android: use an explicit intent:// URL targeting the Gmail
                // package (com.google.android.gm) with a mailto data URI.
                // This opens the Gmail app's compose screen directly — if
                // the app isn't installed, Android's intent system falls
                // back to the Play Store listing for Gmail automatically.
                const ua = navigator.userAgent || '';
                const isAndroid = /android/i.test(ua);
                const isIOS = /iphone|ipad|ipod/i.test(ua);

                if (isAndroid) {
                  e.preventDefault();
                  // The recipient must sit where mailto: puts it — directly
                  // after the scheme, before the "?". Putting it in a "to="
                  // query param (as a generic mailto query string) is NOT
                  // read by Gmail's intent handler and results in a blank
                  // "To" field.
                  const intentUrl = `intent://${encodeURIComponent(to)}?subject=${encodeURIComponent(su)}&body=${encodeURIComponent(bo)}#Intent;scheme=mailto;package=com.google.android.gm;end`;
                  window.location.href = intentUrl;
                } else if (!isIOS) {
                  // Desktop: googlegmail:// scheme isn't registered — go
                  // straight to Gmail web compose instead.
                  e.preventDefault();
                  const webTo = encodeURIComponent(to);
                  const webSu = encodeURIComponent(su);
                  const webBo = encodeURIComponent(bo);
                  window.open(`https://mail.google.com/mail/?view=cm&fs=1&to=${webTo}&su=${webSu}&body=${webBo}`, '_blank', 'noopener,noreferrer');
                }
                // iOS: let the default href (googlegmail://co?...) fire —
                // this is the scheme the Gmail app registers on iOS.
              }}
              className="flex items-center justify-center gap-2 px-4 py-3 font-display text-xs tracking-[0.2em] hover:opacity-90 no-underline border border w-full"
              style={{ background: '#FCF7F2', color: '#002855', fontWeight: 700, textDecoration: 'none' }}
            >
              <Mail className="w-4 h-4" /> GMAIL APP
            </a>
            <p className="text-[10px] italic ocean mt-2">Opens a new draft in the Gmail app on your phone, pre-filled with the recipients, subject and order details. On desktop this opens Gmail in your browser.</p>
          </div>
          {client?.email && !isBlockedEmail(client.email) && !includesClient && (
            <p className="text-[11px] italic" style={{ color: '#BC8D26' }}>
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
      style={{ color: '#002855', fontWeight: 700 }}
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
    channel: 'B2B',
    leadSource: 'Cold Call',
    priority: 'Medium',
    accountManager: defaultRep || '',
    status: 'New',
    notes: '',
    paymentTerms: 'COD',
    prospectedAmount: 0,
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  const [showCloseConfirm, setShowCloseConfirm] = useState(false);

  const newIsDirty = !!form.venue.trim() || !!form.firstName.trim() || !!form.email.trim();
  const handleNewClose = () => { if (newIsDirty) { setShowCloseConfirm(true); } else { onClose(); } };

  const setF = (k, v) => { setForm(f => ({ ...f, [k]: v })); setError(''); };

  const handleCreate = async () => {
    // Validate required fields
    if (!form.venue.trim()) { setError('Venue name is required.'); return; }
    if (!form.channel) { setError('Channel is required.'); return; }
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
    <div style={{ position:"fixed", inset:0, zIndex:60, display:"flex", alignItems:"center", justifyContent:"center", padding:16, background:"rgba(0,40,85,0.82)", backdropFilter:"blur(4px)", overflowY:"auto" }} onClick={handleNewClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background:"#FCF7F2", width:"100%", maxWidth:672, maxHeight:"92vh", overflowY:"auto", border:"2px solid #BC8D26" }}>

        {/* Unsaved warning */}
        {showCloseConfirm && (
          <div style={{ position:'fixed', inset:0, zIndex:200, display:'flex', alignItems:'center', justifyContent:'center', background:'rgba(0,40,85,0.85)', padding:20 }} onClick={e => e.stopPropagation()}>
            <div className="premium-card" style={{ maxWidth:320, width:'100%', padding:28, textAlign:'center' }}>
              <p className="font-display" style={{ fontSize:11, letterSpacing:'0.3em', color:'#BC8D26', fontWeight:600, marginBottom:8 }}>DISCARD CLIENT?</p>
              <p style={{ fontSize:13, color:'#002855', marginBottom:20 }}>You have unsaved data in this form. Are you sure you want to close?</p>
              <div style={{ display:'flex', gap:8 }}>
                <button onClick={() => setShowCloseConfirm(false)}
                  style={{ flex:1, padding:'10px', border:'1px solid rgba(0,40,85,0.2)', background:'none', fontFamily:"'Cinzel',serif", fontSize:9, letterSpacing:'0.2em', color:'#002855', cursor:'pointer', fontWeight:600 }}>
                  KEEP EDITING
                </button>
                <button onClick={() => { setShowCloseConfirm(false); onClose(); }}
                  style={{ flex:1, padding:'10px', background:'#CC233A', border:'none', fontFamily:"'Cinzel',serif", fontSize:9, letterSpacing:'0.2em', color:'#FCF7F2', cursor:'pointer', fontWeight:700 }}>
                  DISCARD
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Header */}
        <div className="p-4 flex items-center justify-between sticky top-0 z-10" style={{ background: '#002855' }}>
          <div className="flex items-center gap-3">
            <UserPlus className="w-5 h-5 flex-shrink-0" style={{ color: '#DBB85E' }} />
            <div>
              <p className="font-display text-[9px] tracking-[0.3em]" style={{ color: '#DBB85E', fontWeight: 600 }}>NEW CLIENT</p>
              <h2 className="font-display text-lg" style={{ color: '#FCF7F2', fontWeight: 700, letterSpacing: '0.06em' }}>
                {defaultRep ? `ADD TO ${defaultRep.toUpperCase()}'S BOOK` : 'ADD NEW CLIENT'}
              </h2>
            </div>
          </div>
          <button type="button" onClick={handleNewClose} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'rgba(255,255,255,0.7)', padding: '4px', lineHeight: 1, fontSize: 22, fontWeight: 300 }}>✕</button>
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
              style={{ borderColor: !form.venue.trim() && error ? '#CC233A' : 'rgba(0,40,85,0.2)' }}
              autoFocus
            />
          </div>

          {/* Channel — critical for visit logging */}
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>CHANNEL * <span className="italic normal-case text-[9px] ocean">(required for visit logging)</span></label>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
              {CHANNELS.map(ch => (
                <button key={ch} type="button"
                  onClick={() => setF('channel', ch)}
                  className={`py-3 font-display text-xs tracking-[0.15em] border-2 transition-all ${form.channel === ch ? 'bg-ink border-ink' : 'bg-cream border'}`}
                  style={{ color: form.channel === ch ? '#FCF7F2' : '#002855', fontWeight: 700 }}>
                  {ch.toUpperCase()}
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
            <div>
              <label className="font-display text-[10px] tracking-[0.25em] copper mb-1 block" style={{ fontWeight: 600 }}>LOCATION</label>
              <select value={form.location} onChange={(e) => setF('location', e.target.value)} className="w-full px-3 py-2.5 border border bg-cream font-body text-sm focus:outline-none focus:border-copper">
                <option value="">— Select area —</option>
                {LOCATIONS.map(l => <option key={l}>{l}</option>)}
              </select>
            </div>
            <div>
              <label className="font-display text-[10px] tracking-[0.25em] copper mb-1 block" style={{ fontWeight: 600 }}>PAYMENT TERMS</label>
              <select value={form.paymentTerms} onChange={(e) => setF('paymentTerms', e.target.value)} className="w-full px-3 py-2.5 border border bg-cream font-body text-sm focus:outline-none focus:border-copper">
                {PAYMENT_TERMS.map(t => <option key={t}>{t}</option>)}
              </select>
            </div>
          </div>

          {/* Prospected Amount — B2B only */}
          {form.channel === 'B2B' && (
            <div>
              <label className="font-display text-[10px] tracking-[0.25em] copper mb-1 block" style={{ fontWeight: 600 }}>PROSPECTED AMOUNT (R)</label>
              <p className="italic ocean mb-2" style={{ fontSize: 10 }}>Estimated value of this B2B opportunity in Rands</p>
              <input
                type="number"
                min="0"
                step="100"
                value={form.prospectedAmount || ''}
                onChange={(e) => setF('prospectedAmount', Number(e.target.value) || 0)}
                placeholder="e.g. 50000"
                className="w-full px-3 py-2.5 border border bg-cream font-body text-sm focus:outline-none focus:border-copper"
              />
            </div>
          )}

          {/* Account manager — full width */}
          <div>
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>ACCOUNT MANAGER *</label>
            <div className="grid grid-cols-3 md:grid-cols-5 gap-2">
              {SALES_REPS.map(r => (
                <button key={r} type="button"
                  onClick={() => setF('accountManager', r)}
                  className={`py-2.5 font-display text-xs tracking-[0.15em] border transition-all ${form.accountManager === r ? 'bg-ink border-ink' : 'bg-cream border'}`}
                  style={{ color: form.accountManager === r ? '#FCF7F2' : '#002855', fontWeight: 700 }}>
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
            <div className="px-3 py-3 border-l-4 text-xs" style={{ borderLeftColor: '#CC233A', background: 'rgba(204,35,58,0.08)', color: '#CC233A' }}>
              <span className="font-display tracking-wider" style={{ fontWeight: 700 }}>ERROR: </span>{error}
            </div>
          )}

          {/* Confirmation message */}
          <p className="text-[10px] italic ocean">
            This client will be saved to Supabase and immediately available for visit logging.
          </p>

          {/* Actions */}
          <div className="flex gap-3 pt-3 border-t">
            <button type="button" onClick={onClose} className="flex-none px-4 py-3 font-display text-xs tracking-[0.2em] ink border border" style={{ fontWeight: 700 }}>
              CANCEL
            </button>
            <button
              type="button"
              onClick={handleCreate}
              disabled={saving}
              className={`flex-1 flex items-center justify-center gap-2 py-3 font-display text-xs tracking-[0.2em] transition-all ${saving ? 'opacity-60 cursor-not-allowed' : ''}`}
              style={{ background: '#BC8D26', color: '#FCF7F2', fontWeight: 700 }}
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
function ClientDetailModal({ client, visits, onClose, onUpdate, onPlaceOrder, onDelete, onDeleteVisit }) {
  const [edit, setEdit] = useState(false);
  const [form, setForm] = useState(client);
  const [showUnsaved, setShowUnsaved] = useState(false);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [tagsSaved, setTagsSaved] = useState(false);

  const isDirty = JSON.stringify(form) !== JSON.stringify(client);

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

  const handleClose = () => {
    if (isDirty) { setShowUnsaved(true); }
    else { onClose(); }
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
    <div style={{ position:"fixed", inset:0, zIndex:50, display:"flex", alignItems:"center", justifyContent:"center", padding:16, background:"rgba(0,40,85,0.82)", backdropFilter:"blur(4px)", overflowY:"auto" }} onClick={handleClose}>
      <div style={{ background:"#FCF7F2", width:"100%", maxWidth:580, margin:"16px 0", border:"2px solid #BC8D26", maxHeight:"95vh", overflowY:"auto" }} onClick={(e) => e.stopPropagation()}>

        {/* Unsaved changes warning */}
        {showUnsaved && (
          <div style={{ position:'fixed', inset:0, zIndex:200, display:'flex', alignItems:'center', justifyContent:'center', background:'rgba(0,40,85,0.85)', padding:20 }}>
            <div className="premium-card" style={{ maxWidth:340, width:'100%', padding:28, textAlign:'center' }}>
              <p className="font-display" style={{ fontSize:11, letterSpacing:'0.3em', color:'#BC8D26', fontWeight:600, marginBottom:8 }}>UNSAVED CHANGES</p>
              <p style={{ fontSize:13, color:'#002855', marginBottom:20 }}>You have unsaved edits on <strong>{client.venue}</strong>. Save before leaving?</p>
              <div style={{ display:'flex', gap:8 }}>
                <button onClick={() => { setShowUnsaved(false); onClose(); }}
                  style={{ flex:1, padding:'10px', border:'1px solid rgba(0,40,85,0.2)', background:'none', fontFamily:"'Cinzel',serif", fontSize:9, letterSpacing:'0.2em', color:'#002855', cursor:'pointer', fontWeight:600 }}>
                  DISCARD
                </button>
                <button onClick={() => { save(); setShowUnsaved(false); onClose(); }}
                  style={{ flex:1, padding:'10px', background:'#BC8D26', border:'none', fontFamily:"'Cinzel',serif", fontSize:9, letterSpacing:'0.2em', color:'#FCF7F2', cursor:'pointer', fontWeight:700 }}>
                  SAVE & CLOSE
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Delete confirmation */}
        {showDeleteConfirm && (
          <div style={{ position:'fixed', inset:0, zIndex:200, display:'flex', alignItems:'center', justifyContent:'center', background:'rgba(0,40,85,0.85)', padding:20 }}>
            <div className="premium-card" style={{ maxWidth:360, width:'100%', padding:28, textAlign:'center' }}>
              <p className="font-display" style={{ fontSize:11, letterSpacing:'0.3em', color:'#CC233A', fontWeight:600, marginBottom:8 }}>DELETE CLIENT</p>
              <p style={{ fontSize:14, color:'#002855', marginBottom:6 }}>Remove <strong>{client.venue}</strong>?</p>
              <p style={{ fontSize:11, color:'#CC233A', fontStyle:'italic', marginBottom:20 }}>This cannot be undone.</p>
              <div style={{ display:'flex', gap:8 }}>
                <button onClick={() => setShowDeleteConfirm(false)}
                  style={{ flex:1, padding:'11px', border:'1px solid rgba(0,40,85,0.2)', background:'none', fontFamily:"'Cinzel',serif", fontSize:9, letterSpacing:'0.2em', color:'#002855', cursor:'pointer', fontWeight:600 }}>
                  CANCEL
                </button>
                <button onClick={async () => { setShowDeleteConfirm(false); await onDelete(); }}
                  style={{ flex:1, padding:'11px', background:'#CC233A', border:'none', fontFamily:"'Cinzel',serif", fontSize:9, letterSpacing:'0.2em', color:'#FCF7F2', cursor:'pointer', fontWeight:700 }}>
                  DELETE
                </button>
              </div>
            </div>
          </div>
        )}

        <div className="p-4 md:p-6 relative overflow-hidden" style={{ background: '#002855' }}>
          <RaysBackdrop opacity={0.06} />
          <div className="flex items-start justify-between gap-4 relative">
            <div className="flex items-start gap-3">
              <AvanteLogo height={36} />
              <div>
                <p className="font-display text-[9px] tracking-[0.3em]" style={{ color: '#DBB85E', fontWeight: 600 }}>{client.channel?.toUpperCase() || 'NO CHANNEL'}</p>
                <h2 className="font-display text-xl md:text-3xl mt-0.5" style={{ color: '#FCF7F2', fontWeight: 700, letterSpacing: '0.05em' }}>{client.venue}</h2>
                <p className="italic text-xs mt-0.5" style={{ color: '#DBB85E' }}>{client.location || 'No location'}</p>
              </div>
            </div>
            <div style={{ display:'flex', alignItems:'center', gap:8, flexShrink:0 }}>
              {onPlaceOrder && (
                <button type="button" onClick={() => onPlaceOrder(client)}
                  style={{ padding:'7px 14px', background:'#BC8D26', border:'none', color:'#FCF7F2', fontFamily:"'Cinzel',serif", fontSize:9, letterSpacing:'0.2em', fontWeight:700, cursor:'pointer', display:'flex', alignItems:'center', gap:6 }}>
                  <span style={{ fontSize:13, fontWeight:300 }}>+</span> PLACE ORDER
                </button>
              )}
              {onDelete && (
                <button type="button" onClick={() => setShowDeleteConfirm(true)}
                  style={{ padding:'7px 10px', background:'none', border:'1px solid rgba(255,255,255,0.25)', color:'rgba(255,100,100,0.8)', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center' }}
                  title="Delete client">
                  <Trash2 style={{ width:14, height:14 }} />
                </button>
              )}
              <button type="button" onClick={handleClose} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'rgba(255,255,255,0.65)', padding: '4px 8px', fontSize: 22, fontWeight: 200, lineHeight: 1 }}>✕</button>
            </div>
          </div>
        </div>

        <div className="p-4 md:p-6 space-y-5">
          {/* Quick stats — 2 col on mobile, 4 on desktop */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 pb-4 border-b">
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
              <p style={{ fontSize: 10, color: '#5A7A99', fontStyle: 'italic', marginTop: 2 }}>Last contact</p>
              <p className="font-display text-base ink mt-1" style={{ fontWeight: 700 }}>
                {client.lastContacted ? client.lastContacted + (daysSinceLast !== null ? ` (${daysSinceLast} days ago)` : '') : '—'}
              </p>
            </div>
            <div>
              <p className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>PAYMENT TERMS</p>
              <p className="font-display text-base ink mt-1" style={{ fontWeight: 700 }}>{client.paymentTerms || 'COD'}</p>
            </div>
          </div>

          {/* Status quick-change — separate row */}
          <div className="pb-3 border-b">
            <p className="font-display text-[10px] tracking-[0.3em] copper mb-2" style={{ fontWeight: 600 }}>STATUS</p>
            <select
              value={form.status || client.status || ''}
              onChange={async (e) => {
                const newStatus = e.target.value;
                setForm(f => ({ ...f, status: newStatus }));
                await onUpdate({ status: newStatus });
              }}
              style={{ padding: '6px 10px', border: '1px solid rgba(0,40,85,0.2)', background: '#FCF7F2', fontFamily: "'Cinzel', serif", fontSize: 11, letterSpacing: '0.1em', fontWeight: 700, color: '#002855', cursor: 'pointer', outline: 'none' }}
            >
              {(() => {
                const baseOptions = getStatusesForChannel(form.channel || client.channel);
                const currentStatus = form.status || client.status;
                // Always include the current status even if not in the list
                const allOptions = currentStatus && !baseOptions.includes(currentStatus)
                  ? [currentStatus, ...baseOptions]
                  : baseOptions;
                return allOptions.map(s => (
                  <option key={s} value={s}>{s.toUpperCase()}</option>
                ));
              })()}
            </select>
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
              <SelectField label="Status" value={form.status} edit={edit} onChange={(v) => setForm({ ...form, status: v })} options={getStatusesForChannel(form.channel || client.channel)} />
              <SelectField label="Channel" value={form.channel} edit={edit} onChange={(v) => setForm({ ...form, channel: v })} options={CHANNELS} />
              <SelectField label="Location" value={form.location} edit={edit} onChange={(v) => setForm({ ...form, location: v })} options={['', ...LOCATIONS]} />
              <SelectField label="Payment Terms" value={form.paymentTerms} edit={edit} onChange={(v) => setForm({ ...form, paymentTerms: v })} options={PAYMENT_TERMS} />
              {/* Prospected Amount — B2B channel only */}
              {(form.channel || client.channel) === 'B2B' && (
                <div>
                  <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>PROSPECTED AMOUNT (R)</label>
                  {edit ? (
                    <input
                      type="number"
                      min="0"
                      step="100"
                      value={form.prospectedAmount || ''}
                      onChange={(e) => setForm({ ...form, prospectedAmount: Number(e.target.value) || 0 })}
                      placeholder="e.g. 50000"
                      className="w-full px-3 py-2 border border bg-cream font-body text-sm focus:outline-none focus:border-copper"
                    />
                  ) : (
                    <div className="text-sm ink py-1">
                      {Number(form.prospectedAmount) > 0 ? ZAR(form.prospectedAmount) : <span className="italic ocean">—</span>}
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>

          {(edit || client.notes) && (
            <div>
              <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>NOTES</label>
              {edit ? (
                <textarea value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} rows="3" className="w-full px-3 py-2 border border bg-cream font-body text-sm focus:outline-none focus:border-copper resize-none" />
              ) : (
                <p className="text-sm ink italic">{client.notes}</p>
              )}
            </div>
          )}

          {/* === TAG AGENTS === */}
          <div className="pt-4 border-t">
            <label className="font-display text-[10px] tracking-[0.3em] copper mb-2 block" style={{ fontWeight: 600 }}>TAG AGENTS</label>
            <p style={{ fontSize: 10, color: '#5A7A99', fontStyle: 'italic', marginBottom: 10 }}>Tag team members who should follow up or are involved with this client.</p>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              {SALES_REPS.map(rep => {
                const currentTags = form.clientTags || [];
                const tagged = currentTags.includes(rep);
                return (
                  <button key={rep} type="button"
                    onClick={() => {
                      const next = tagged
                        ? currentTags.filter(r => r !== rep)
                        : [...currentTags, rep];
                      setForm(f => ({ ...f, clientTags: next }));
                      onUpdate({ clientTags: next });
                    }}
                    style={{ padding: '8px 16px', fontFamily: "'Cinzel',serif", fontSize: 9, letterSpacing: '0.2em', fontWeight: 700, border: '1px solid', borderColor: tagged ? '#BC8D26' : 'rgba(0,40,85,0.2)', background: tagged ? '#BC8D26' : 'transparent', color: tagged ? '#FCF7F2' : '#002855', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, transition: 'all 0.15s' }}>
                    {tagged && <span style={{ fontSize: 10 }}>✓</span>} @{rep.toUpperCase()}
                  </button>
                );
              })}
            </div>
            {(form.clientTags || []).length > 0 && (
              <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                {(form.clientTags || []).map(r => (
                  <span key={r} style={{ padding: '3px 10px', background: 'rgba(188,141,38,0.15)', color: '#BC8D26', fontFamily: "'Cinzel',serif", fontSize: 8, letterSpacing: '0.15em', fontWeight: 700 }}>@{r.toUpperCase()} TAGGED</span>
                ))}
              </div>
            )}
          </div>

          {/* === PURCHASE HISTORY (SKUs aggregated) === */}
          {skuTotals.length > 0 && (
            <div className="pt-5 border-t">
              <div className="flex items-center justify-between mb-3">
                <p className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>PURCHASE HISTORY · SKU BREAKDOWN</p>
                <span className="text-[10px] italic ocean">All-time totals across {visits.filter(v => v.items && v.items.length > 0).length} order{visits.filter(v => v.items && v.items.length > 0).length === 1 ? '' : 's'}</span>
              </div>
              <div className="border border">
                <div className="grid grid-cols-12 gap-2 px-3 py-2  border-b">
                  <div className="col-span-6 font-display text-[10px] tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>SKU</div>
                  <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-center" style={{ fontWeight: 600 }}>UNITS</div>
                  <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-right" style={{ fontWeight: 600 }}>REVENUE</div>
                  <div className="col-span-2 font-display text-[10px] tracking-[0.2em] ocean text-right" style={{ fontWeight: 600 }}>LAST</div>
                </div>
                {skuTotals.map(s => (
                  <div key={s.skuId} className="grid grid-cols-12 gap-2 px-3 py-2 border-b items-center text-xs">
                    <div className="col-span-6 ink font-display" style={{ fontWeight: 700 }}>{s.name}</div>
                    <div className="col-span-2 text-center ink">{s.qty}</div>
                    <div className="col-span-2 text-right copper font-display" style={{ fontWeight: 700 }}>{ZAR(s.revenue)}</div>
                    <div className="col-span-2 text-right ocean italic">{s.lastDate || '—'}</div>
                  </div>
                ))}
                <div className="grid grid-cols-12 gap-2 px-3 py-2.5" style={{ background: '#002855' }}>
                  <div className="col-span-6 font-display text-[10px] tracking-[0.3em]" style={{ color: '#DBB85E', fontWeight: 600 }}>LIFETIME TOTAL</div>
                  <div className="col-span-2 text-center font-display" style={{ color: '#FCF7F2', fontWeight: 700 }}>{skuTotals.reduce((s, x) => s + x.qty, 0)}</div>
                  <div className="col-span-4 text-right font-display text-base" style={{ color: '#FCF7F2', fontWeight: 700 }}>{ZAR(skuTotals.reduce((s, x) => s + x.revenue, 0))}</div>
                </div>
              </div>
            </div>
          )}

          {/* === CALL CYCLE / VISIT TIMELINE === */}
          {orderedVisits.length > 0 ? (
            <div className="pt-5 border-t">
              <div className="flex items-center justify-between mb-3">
                <p className="font-display text-[10px] tracking-[0.3em] copper" style={{ fontWeight: 600 }}>CALL CYCLE · VISIT TIMELINE ({orderedVisits.length})</p>
                <span className="text-[10px] italic ocean">Newest first</span>
              </div>
              <div className="space-y-3">
                {orderedVisits.map((v, idx) => {
                  const isHistorical = v.isHistorical;
                  // Historical entries get a muted ocean blue; logged visits get outcome-based color
                  const outcomeColor = isHistorical
                    ? '#5A7A99'
                    : v.outcome === 'Sold In' ? '#2d8659'
                    : v.outcome === 'Rejected' ? '#CC233A'
                    : '#BC8D26';
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
                        <div className="flex items-baseline justify-between gap-2 px-3 py-2 border-b" style={{ background: isLatest ? 'rgba(219,184,94,0.08)' : isHistorical ? 'rgba(90,122,153,0.04)' : 'transparent' }}>
                          <div className="flex items-baseline gap-3 flex-wrap">
                            <span className="font-display ink text-sm tracking-wide" style={{ fontWeight: 700 }}>{v.date || 'No date'}</span>
                            <span className="text-[10px] font-display tracking-[0.2em] ocean" style={{ fontWeight: 600 }}>BY {v.salesRep?.toUpperCase() || '—'}</span>
                            {v.contactMethod && <span className="text-[9px] italic ocean">via {v.contactMethod}</span>}
                            {isLatest && <span className="text-[9px] font-display tracking-[0.2em] px-1.5 py-0.5" style={{ background: '#DBB85E', color: '#002855', fontWeight: 700 }}>LATEST</span>}
                            {isHistorical && <span className="text-[9px] font-display tracking-[0.2em] px-1.5 py-0.5 border" style={{ borderColor: '#5A7A99', color: '#5A7A99', fontWeight: 700 }}>IMPORTED</span>}
                          </div>
                          <div className="flex items-center gap-2">
                            <span className="text-[10px] font-display tracking-wider px-2 py-0.5" style={{ background: outcomeColor, color: '#FCF7F2', fontWeight: 600 }}>
                              {(v.outcome || 'VISIT').toUpperCase()}
                            </span>
                            {v.saleAmount > 0 && <span className="font-display text-sm copper" style={{ fontWeight: 700 }}>{ZAR(v.saleAmount)}</span>}
                            {onDeleteVisit && (
                              <button type="button"
                                onClick={() => onDeleteVisit(v.id)}
                                style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '2px 4px', color: 'rgba(0,40,85,0.25)', flexShrink: 0, lineHeight: 1 }}
                                onMouseEnter={e => e.currentTarget.style.color = '#CC233A'}
                                onMouseLeave={e => e.currentTarget.style.color = 'rgba(0,40,85,0.25)'}
                                title="Delete this visit">
                                <Trash2 style={{ width: 12, height: 12 }} />
                              </button>
                            )}
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
                          {/* Tagged agents */}
                          {v.taggedReps && v.taggedReps.length > 0 && (
                            <div>
                              <p className="font-display text-[9px] tracking-[0.2em] copper mb-0.5" style={{ fontWeight: 600 }}>TAGGED</p>
                              <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                                {v.taggedReps.map(r => (
                                  <span key={r} style={{ padding: '2px 8px', background: 'rgba(188,141,38,0.15)', color: '#BC8D26', fontFamily: "'Cinzel',serif", fontSize: 8, letterSpacing: '0.1em', fontWeight: 700 }}>@{r.toUpperCase()}</span>
                                ))}
                              </div>
                            </div>
                          )}
                          {/* Follow-up date */}
                          {v.followUpDate && (
                            <div>
                              <p className="font-display text-[9px] tracking-[0.2em] copper mb-0.5" style={{ fontWeight: 600 }}>FOLLOW-UP DUE</p>
                              <p className="text-xs ink" style={{ fontWeight: 700 }}>{v.followUpDate}</p>
                            </div>
                          )}
                          {/* Follow-up notes */}
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
            <div className="pt-5 border-t">
              <p className="font-display text-[10px] tracking-[0.3em] copper mb-3" style={{ fontWeight: 600 }}>CALL CYCLE · VISIT TIMELINE</p>
              <div className="border border-dashed border p-6 text-center">
                <p className="text-xs italic ocean">No visits logged for this client yet.</p>
                <p className="text-xs italic ocean mt-1">Use "Log Visit" to start tracking the call cycle.</p>
              </div>
            </div>
          )}

          <div className="flex items-center justify-end gap-3 pt-4 border-t">
            {edit ? (
              <>
                <button type="button" onClick={() => { setForm(client); setEdit(false); }} className="px-5 py-2.5 font-display text-xs tracking-[0.25em] ink" style={{ fontWeight: 700 }}>CANCEL</button>
                <button type="button" onClick={save} className="bg-copper hover:bg-gold px-6 py-2.5 font-display text-xs tracking-[0.25em]" style={{ color: '#FCF7F2', fontWeight: 700 }}>SAVE CHANGES</button>
              </>
            ) : (
              <>
                <button type="button"
                  onClick={async () => {
                    const prevTags = client.clientTags || [];
                    const newTags = form.clientTags || [];
                    notifyNewTags({
                      prevTags, newTags,
                      clientName: client.venue,
                      location: client.location,
                      channel: client.channel,
                      taggedBy: client.accountManager,
                      type: 'client',
                      status: client.status,
                      notes: client.notes,
                    });
                    await onUpdate({ clientTags: newTags });
                    setTagsSaved(true);
                    setTimeout(() => setTagsSaved(false), 2000);
                  }}
                  className="px-6 py-2.5 font-display text-xs tracking-[0.25em]"
                  style={{ background: tagsSaved ? '#2d8659' : 'transparent', color: tagsSaved ? '#FCF7F2' : '#002855', border: '1px solid', borderColor: tagsSaved ? '#2d8659' : 'rgba(0,40,85,0.25)', fontWeight: 700, display: 'flex', alignItems: 'center', gap: 6 }}>
                  <Save style={{ width: 13, height: 13 }} /> {tagsSaved ? 'TAGS SAVED' : 'SAVE TAGS'}
                </button>
                <button type="button" onClick={() => setEdit(true)} className="bg-ink hover:bg-copper px-6 py-2.5 font-display text-xs tracking-[0.25em]" style={{ color: '#FCF7F2', fontWeight: 700 }}>EDIT CLIENT</button>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

function Field({ label, value, edit, onChange, icon: Icon }) {
  if (!edit && !value) return null;
  return (
    <div>
      <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>{label.toUpperCase()}</label>
      {edit ? (
        <input type="text" value={value || ''} onChange={(e) => onChange(e.target.value)} className="w-full px-3 py-2 border border bg-cream font-body text-sm focus:outline-none focus:border-copper" />
      ) : (
        <div className="flex items-center gap-2 text-sm ink py-1">
          {Icon && <Icon className="w-3.5 h-3.5 ocean" />}
          {value}
        </div>
      )}
    </div>
  );
}

function SelectField({ label, value, edit, onChange, options }) {
  if (!edit && !value) return null;
  // Always include the current value as an option even if not in the list
  const allOptions = value && !options.includes(value) ? [value, ...options] : options;
  return (
    <div>
      <label className="font-display text-[10px] tracking-[0.3em] copper mb-1 block" style={{ fontWeight: 600 }}>{label.toUpperCase()}</label>
      {edit ? (
        <select value={value || ''} onChange={(e) => onChange(e.target.value)} className="w-full px-3 py-2 border border bg-cream font-body text-sm focus:outline-none focus:border-copper">
          {allOptions.map(o => <option key={o} value={o}>{o}</option>)}
        </select>
      ) : (
        <div className="text-sm ink py-1">{value}</div>
      )}
    </div>
  );
}
