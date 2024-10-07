async function pubChem() {
  var range = SpreadsheetApp.getActiveSheet().getActiveCell();

  if (range.getColumn() !== 1) { // never run unless on column 1 // i dont need to set row because it doesn't run when the col 44 has values
    // Logger.log(range.getColumn())
    // Logger.log(range.getRowIndex())
    range = range.offset(0,-(range.getColumn()-1)).activate(); 
  }

  // BEGIN INTERATION OF SHEET
  while (range.getRowIndex() <= SpreadsheetApp.getActiveSheet().getLastRow() && !range.offset(0,44).getValue()) { // check that not end of sheet and not already run on

    // init variables:
    var name = range.offset(0,1).getValue().trim().replace(' ', '%20'); // percentile encode spaces
    var rn = range.offset(0,4).getValue().trim();

    // synonyms, chemical forumula, flash point, density weight, melting point, boiling point, molecular weight, pH value, unit, physical state, compatability category, special hazard, UN number, UN pack group, TDG primary, TDG secondary, TDG teritary, storage requirements, GHS classification, GHS hazard statements, GHS signal word, GHS precautionary statement codes, GHS pictogram(s), NFPA Diamond, DOT Guide  
    var syn = '', cf = '', fp = '', dw = '', mp = '', bp = '', mw = '', pH = '', unit = '', ps = '', cs = '', cc = '', sh = '', un_no = '', un_pg = '', tdg1 = '', tdg2 = '', tdg3 = '', strReq = '', ghs_class = '', ghs_hs = '', ghs_psc = '', ghs_signal = '', ghs_pic = '', nfpa_dia = '', dot_guide = '', erg_hc = '';

    var flag_rn = false; // bad cas entry
    var tdg_flag = false; // tdg failure, attempt to find and convert dot 

    const headers =['Depositor-Supplied%20Synonyms','Boiling%20Point','Melting%20Point','Flash%20Point','Density','Physical%20Description','Chemical%20Classes', 'Reactive%20Group','Hazards%20Summary','UN%20Number','UN%20Classification','GHS%20Classification','Hazard%20Classes%20and%20Categories','NFPA%20Hazard%20Classification', 'DOT%20Label', 'DOT%20Emergency%20Guidelines','Hazards%20Identification'];

    // BEGIN ITERATION OF ROW (INIT HEADERS)
    try { 
      Logger.log('ATTEMPTING: '+name+', '+rn)
      if (rn && !rn.includes('Z')) { // real cas
        flag_rn = true;
        var urlCompound = 'https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/'+rn+'/property/MolecularFormula,MolecularWeight,IUPACName,Title,CanonicalSMILES/json';
      } else if (name) { // chematix generated cas - try name
        range.offset(0,3).setValue('true');
        var urlCompound = 'https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/'+name.split('.')[0]+'/property/MolecularFormula,MolecularWeight,IUPACName,Title,CanonicalSMILES/json'; // TODO: for now dirty fix for naming conventions with formulas
      }

      // fetch request and parse data returned
      const responseCompound = UrlFetchApp.fetch(urlCompound);
      var dataCompound = JSON.parse(responseCompound.getContentText());  
      // Logger.log(dataCompound);
      cid = (dataCompound.PropertyTable.Properties[0].CID).toString().split('.')[0]
      cf = dataCompound.PropertyTable.Properties[0].MolecularFormula;
      mw = dataCompound.PropertyTable.Properties[0].MolecularWeight;
      cs = dataCompound.PropertyTable.Properties[0].CanonicalSMILES;
      name = dataCompound.PropertyTable.Properties[0].Title; // sanitize naming convention w/ pubchem title
      if (dataCompound.PropertyTable.Properties[0].IUPACName && dataCompound.PropertyTable.Properties[0].IUPACName.toLowerCase() !== name.toLowerCase()) { // check that the title is not equal to IUPAC for synonym
        syn = dataCompound.PropertyTable.Properties[0].IUPACName + ';'; 
      }

      

      if (!flag_rn) { // if chematix cas and positive fetch return - get real cas
        var urlDataRN = 'https://pubchem.ncbi.nlm.nih.gov/rest/pug_view/data/compound/'+cid+'/JSON/?response_type=display&heading=CAS'
        try {
          var responseDataRN = UrlFetchApp.fetch(urlDataRN);
          var dataSiteRN = JSON.parse(responseDataRN.getContentText());
          rn = dataSiteRN.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[0].String
        }
        catch (err){
          Logger.log('rn format issue:' +err+'\n')
        }
      }

      // BEGIN ITERATION OF HEADERS
      for (let i = 0; i < headers.length; i++) {
        var urlDataCompound = 'https://pubchem.ncbi.nlm.nih.gov/rest/pug_view/data/compound/'+cid+'/JSON/?response_type=display&heading='+headers[i];  
        try {
          var responseCompoundData = UrlFetchApp.fetch(urlDataCompound);
          var dataSiteCompound = JSON.parse(responseCompoundData.getContentText());

           if (dataSiteCompound) {
            Logger.log('Found: '+headers[i]+'\n');

            // SYNOMYMS
            if (headers[i] == 'Depositor-Supplied%20Synonyms') { 
              for (let j = 0; j < 7 && j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup).length; j++) { // cap at 6 synonyms or less

                if (rn !== dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[j].String && !syn.includes(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[j].String.toLowerCase()) && name.toLowerCase() !== dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[j].String.toLowerCase()) {

                  syn += dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[j].String + ';' 
                } 
              }
              Logger.log('synonyms:'+syn);     
            } 

            // BOILING POINT
            else if (headers[i] == 'Boiling%20Point') {
              var ref_no = ''
              for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                try {
                  // Logger.log('current bp: '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String)

                  if (dataSiteCompound.Record.Reference) {
                    for (let j = 0; j < Object.keys(dataSiteCompound.Record.Reference).length; j++) {
                      if (dataSiteCompound.Record.Reference[j].SourceName == "ILO-WHO International Chemical Safety Cards (ICSCs)") {
                        ref_no = dataSiteCompound.Record.Reference[j].ReferenceNumber
                        Logger.log('ref no found: '+ref_no)
                      }
                    }
                  }

                  if (ref_no) {
                    for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                      if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].ReferenceNumber == ref_no) {
                        bp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                        break
                      }
                    }
                  }

                  else {
                    if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Description == 'PEER REVIEWED' && dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('°C') && dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('760 mm Hg')) {
                      bp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                      break
                    } 
                    else if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('°C') && dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('760 mm Hg')){
                      bp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                    }            
                    else {
                      if (!bp.includes('°C'))
                      bp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                    }
                  }
                }

                catch (err) { 
                  Logger.log('bp format issue:' +err+'\n')
                  // Logger.log('current bp: '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Number[0]+' '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Unit)

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Number[0] && dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Unit.includes('°C')) {
                    bp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Number[0]+' '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Unit
                    break
                  } 
                }
              }
              Logger.log('boiling point: '+bp)
            } 

            // MELTING POINT 
            else if (headers[i] == 'Melting%20Point') {
              var ref_no = ''
              for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                try {
                  // Logger.log('current mp: '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String)

                  if (dataSiteCompound.Record.Reference) {
                    for (let j = 0; j < Object.keys(dataSiteCompound.Record.Reference).length; j++) {
                      if (dataSiteCompound.Record.Reference[j].SourceName == "ILO-WHO International Chemical Safety Cards (ICSCs)") {
                        ref_no = dataSiteCompound.Record.Reference[j].ReferenceNumber
                        Logger.log('ref no: '+ref_no)
                      }
                    }
                  }

                  if (ref_no) {
                    for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                      if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].ReferenceNumber == ref_no) {
                        mp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                        break
                      }
                    }
                  }
                  
                  else {
                    if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Description == 'PEER REVIEWED' && dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('°C')) {
                      mp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                      break
                    } 
                    else if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('°C')){
                      mp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                    }            
                    else {
                      mp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                    }
                  }

                }
                catch (err) { 
                  Logger.log('mp format issue:' +err+'\n')
                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Number[0] && dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Unit.includes('°C')) {
                    mp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Number[0]+' '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Unit
                    break
                  } 
                }
              }
              // Logger.log('melting point: '+mp)
            } 

            // FLASH POINT
            else if (headers[i] == 'Flash%20Point') {
              var ref_no = ''
              for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                try {
                  // Logger.log('current fp: '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String)
                  if (dataSiteCompound.Record.Reference) {
                    for (let j = 0; j < Object.keys(dataSiteCompound.Record.Reference).length; j++) {
                      if (dataSiteCompound.Record.Reference[j].SourceName == "ILO-WHO International Chemical Safety Cards (ICSCs)") {
                        ref_no = dataSiteCompound.Record.Reference[j].ReferenceNumber
                        Logger.log('ref no: '+ref_no)
                      }
                    }
                  }

                  if (ref_no) {
                    for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                      if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].ReferenceNumber == ref_no) {
                        fp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                        break
                      }
                    }
                  }

                  else {
                    if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Description == 'PEER REVIEWED' && dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('°C')) {
                      fp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                      break
                    }
                    else if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Description == 'PEER REVIEWED'){
                      fp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                    }  
                    else if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('°C')){
                      fp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                    }            
                    else {
                      if (!fp.includes('°C')) {
                        fp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                      }
                    }
                  }

                }
                catch (err) { 
                  Logger.log('fp format issue:' +err+'\n')
                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Number[0] && dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Unit.includes('°C')) {
                    fp = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Number[0]+' '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.Unit
                    break
                  }                  
                  
                }
              }
              // Logger.log('flash point: '+fp)
            } 

            // WEIGHT DENSITY 
            else if (headers[i] == 'Density') {
              var ref_no = ''
              for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                try {
                  // Logger.log('current dw: '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String)

                  if (dataSiteCompound.Record.Reference) {
                    for (let j = 0; j < Object.keys(dataSiteCompound.Record.Reference).length; j++) {
                      if (dataSiteCompound.Record.Reference[j].SourceName == "ILO-WHO International Chemical Safety Cards (ICSCs)") {
                        ref_no = dataSiteCompound.Record.Reference[j].ReferenceNumber
                        Logger.log('ref no: '+ref_no)
                      }
                    }
                  }

                  if (ref_no) {
                    for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                      if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].ReferenceNumber == ref_no) {
                        dw = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                        break
                      }
                    }
                  }

                  else {
                    if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Description == 'PEER REVIEWED') {
                      dw = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                      break
                    }     
                    else if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('g/')){
                      dw = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                    }  
                    else if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('°C')){
                      dw = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                    }                                                  
                    else {
                      if (!dw.includes('g/') || !dw.includes('°C')) {
                        dw = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                      }
                    }
                  }

                }
                catch (err) { 
                  Logger.log('dw format issue:' +err+'\n')
                }
              }
              // Logger.log('density weight: '+dw)
            } 

            // PHYSICAL STATE
            else if (headers[i] == 'Physical%20Description') {
              let solid_count = 0
              let liquid_count = 0
              let gas_count = 0

              for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                try {
                  // Logger.log('current ps: '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String)
                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('solid') || dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('powder')) {
                    solid_count++
                  }
                  else if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('liquid')){
                    liquid_count++
                  }
                  else if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.includes('gas')){
                    gas_count++
                  }                                         
                }                
                catch (err) { 
                  // Logger.log('ps format issue:' +err+'\n')
                }
              }

              max = Math.max(solid_count,liquid_count,gas_count) // let democracy decide

              if (solid_count == max) {
                ps = 'solid'
              }
              else if (liquid_count == max) {
                ps = 'liquid'
              }
              else if (gas_count == max) { // if all were same it might default to this...
                ps = 'gas'
              }
              // Logger.log('physical state: '+ps)
            } 

            // CHEMICAL COMPATIBILITY (Reactive Group) from: https://cameochemicals.noaa.gov/browse/react
            else if (headers[i] == 'Reactive%20Group') {
              for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                try {
                  // Logger.log('current cc: '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String)
                  if (!cc.toLowerCase().includes(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase())) {
                    cc += dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String+';'
                  }       
                }
                catch (err) { 
                  // Logger.log('reactive group format issue:' +err+'\n')
                }
              }
              // Logger.log('chemical compatibility: '+cc)
            } 

            // SPECIAL HAZARD
            else if (headers[i] == 'Hazards%20Summary') {
                try {
                  sh += dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[0].String
                }
                catch (err) { 
                  // Logger.log('reactive group format issue:' +err+'\n')
                }
              // Logger.log('special hazard: '+sh)
            } 

            // UN NUMBER  
            else if (headers[i] == 'UN%20Number') {
              
              for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                try {
                  // Logger.log('current un no.: '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String)
                  if (!un_no.includes(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.split(' ')[0])) {
                    un_no += dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toUpperCase().replace(';',',')+';'           
                  }
                  else { // if they match
                    // Logger.log(un_no)
                    let temp_list = un_no.split(';')
                    // Logger.log(temp_list)
                    for (let k = 0; k < temp_list.length-1; k++) {
                      if (temp_list[k].split(' ')[0] == dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.split(' ')[0] && temp_list[k].length < dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.length) { // facepalm
                        temp_list[k] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toUpperCase().replace(';',',');
                      }
                    }
                    un_no = temp_list.join(';')
                  }
                  // Logger.log(un_no)
                }
                catch (err) { 
                  Logger.log('un_no format issue:' +err+'\n')
                }
              }
              // Logger.log('un number: '+un_no)
            } 

            // TDG CLASS(ES), UN PACK GROUP
            else if (headers[i] == 'UN%20Classification') { 
                try {
                  data = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[0].String.split(/UN Hazard Class|UN Subsidiary Risks|UN Pack Group/)
                  // Logger.log(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[0].String)
                  // Logger.log('un classification: '+data+'\nsize: '+data.length)

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[0].String.match(/UN Hazard Class/)) { // index = 1
                    // Logger.log('hazard class: '+data[1].replace(/[,;:]/g, ' ').trim())
                    tdg1 = data[1].replace(/[,;:]/g, ' ').trim()
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[0].String.match(/UN Subsidiary Risk/)) { // index = 2
                    // Logger.log('subrisks: '+data[2].replace(/[,;:]/g, ' ').trim())
                    sr = data[2].replace(/[,;:]/g, ' ').trim().match(/\b\d+(\.\d+)?\b/g)
                    if (sr.length > 1) {
                      tdg3 = sr[1]
                    }
                    tdg2 = sr[0]
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[0].String.match(/UN Pack Group/)) { // index = 2 | 3
                    if (data.length == 3) {
                      // Logger.log('pack group: '+data[2].replace(/[,;:]/g, ' ').trim())
                      un_pg = data[2].replace(/[,;:]/g, ' ').trim()
                    }
                    else if (data.length == 4) {
                      // Logger.log('pack group: '+data[3].replace(/[,;:]/g, ' ').trim())
                      un_pg = data[3].replace(/[,;:]/g, ' ').trim()
                    }
                  }
                }
                catch (err) { 
                  Logger.log('(un) transport info format issue:' +err+'\n')
                }
              // Logger.log('tdg class(es): '+tdg1+', '+tdg2+', '+tdg3)
              // Logger.log('un pack group: '+un_pg)             
            } 

            // GHS HAZARDS/PRECAUTIONARY STATEMENTS, SIGNAL, PICTOGRAM(S)
            else if (headers[i] == 'GHS%20Classification') {
              // Logger.log('obj size: '+Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length)
              for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
          
                try {        

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Name == 'Pictogram(s)' && !ghs_pic) { // this is a list
                    for (let k = 0; k < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].Markup).length; k++) {
                      ghs_pic += dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].Markup[k].Extra+';'
                    }
                  }

                  else if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Name == 'Signal' && !ghs_signal) {
                    ghs_signal = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String
                  }

                  else if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Name == 'GHS Hazard Statements' && !ghs_hs) { // this is a list
                    // Logger.log('hazard statements')
                    for (let k = 0; k < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup).length; k++) {
                      ghs_hs += dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[k].String.match(/H\d{3}[a-zA-Z]?/g)+';'
                    }
                  }

                  else if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Name == 'Precautionary Statement Codes' && !ghs_psc) {
                    // Logger.log(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String)
                    ghs_psc = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.replace(/ |and/g, '').split(',').join(';')
                  }

                  else if (ghs_pic && ghs_signal && ghs_hs && ghs_psc) { // complete
                    break;
                  }

                }
                catch (err) { 
                  // Logger.log('ghs format issue:' +err+'\n')
                }
              }
              // Logger.log('ghs hazards, precautionary, picto, signal: '+ghs_hs+ghs_psc+ghs_pic+ghs_signal)
            } 

            // GHS CLASSIFICATION
            else if (headers[i] == 'Hazard%20Classes%20and%20Categories') {
              // Logger.log(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup)
              for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup).length; j++) {
                try {
                  // Logger.log('current ghs_class: '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[j].String) 
                    if (!ghs_class.toLowerCase().includes(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[j].String.replace(/\(\d+(\.\d+)?%\)/g, '').trim().toLowerCase())) {
                      ghs_class += dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[j].String.replace(/\(\d+(\.\d+)?%\)/g, '').trim()+';'
                    } 
                }
                catch (err) { 
                  // Logger.log('ghs_class format issue:' +err+'\n')
                }
              }
              // Logger.log('ghs_class: '+ghs_class)
            } 

            // NFPA DIAMOND
            else if (headers[i] == 'NFPA%20Hazard%20Classification') {
                try {
                  nfpa_dia += dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[0].Markup[0].Extra
                  // Logger.log('current nfpa: '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[0].Value.StringWithMarkup[0].String)
                }
                catch (err) { 
                  // Logger.log('nfpa diamond format issue:' +err+'\n')
                } 
                //  Logger.log('nfpa diamond: '+nfpa_dia)             
            }

            // DOT LABEL // TODO: double check explosives are all covered by regex num search
            else if (headers[i] == 'DOT%20Label' && tdg_flag) {
              Logger.log('attempting to pull DOT Label for: '+cid)
              // Logger.log(dataSiteCompound)

              try {  
                const dot_dict = { // this idea is falling apart we might just make a large if else statement. 
                'flammable gas':'2.1',
                'non-flammable gas':'2.2',
                'poisonous gas':'2.3',
                'flammable liquid':'3',
                'combustible liquid': '3', // quick fix maybe just use liquid?
                'flammable solid':'4.1',
                'spontaneously combustible':'4.2',
                'dangerous when wet':'4.3',
                'oxidizer':'5.1',
                'organic peroxide':'5.2',
                'poison':'6.1',
                'infectious substance':'6.2',
                'radioactive':'7',
                'corrosive':'8',
                'miscellaneous hazardous material':'9', // this one may never get used, but just incase...
                };
                var order = {} // find order of DOT classes to determine hiearchy 
          
                for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) { // need to double check that all explosives (class 1) incl. their numbers 
                  // Logger.log('DOT Label: '+dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String)
                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().match(/\b\d(\.\d[A-Za-z]?)?\b/g)) {
                    str_nums = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().match(/\b\d(\.\d[A-Za-z]?)?\b/g)
                    for (let i = 0; i < str_nums.length; i++) {
                      order[str_nums[i]] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf(str_nums[i])
                      
                    }
                  }
          
                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('flammable gas')) {
                  order[dot_dict['flammable gas']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('flammable gas') 
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('non-flammable gas')) {
                    order[dot_dict['non-flammable gas']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('non-flammable gas')
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('poisonous gas')) {
                    order[dot_dict['poisonous gas']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('poisonous gas')
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('flammable liquid')) { 
                    order[dot_dict['flammable liquid']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('flammable liquid')
                  }

                  // quick fix...
                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('combustible liquid')) { // combustable liquid
                    order[dot_dict['combustible liquid']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('combustible liquid')
                  } 
                  
                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('flammable solid')) {
                    order[dot_dict['flammable solid']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('flammable solid')
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('spontaneously combustible')) {
                    order[dot_dict['spontaneously combustible']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('spontaneously combustible')
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('dangerous when wet')) {
                    order[dot_dict['dangerous when wet']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('dangerous when wet')
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('oxidizer')) {
                    order[dot_dict['oxidizer']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('oxidizer')
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('organic peroxide')) {
                    order[dot_dict['organic peroxide']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('organic peroxide')
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('poison')) {
                    order[dot_dict['poison']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('poison')
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('infectious substance')) {
                    order[dot_dict['infectious substance']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('infectious substance')
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('radioactive')) {
                    order[dot_dict['radioactive']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('radioactive')
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('corrosive')) {
                    order[dot_dict['corrosive']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('corrosive')
                  }

                  if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().includes('miscellaneous hazardous material')) { // either this or 'class 9'
                    order[dot_dict['miscellaneous hazardous material']] = dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.toLowerCase().indexOf('miscellaneous hazardous material')
                  }
                }     
                // Logger.log(order) // maybe check to see if tdg already exists from UN pull then append on if space and if enough, do not use tdg vars then set it to dot1, dot2, dot3
                // i am so sorry for this abomination of a permutation... 
                if (Object.keys(order).length == 1) { 
                  tdg1 = Object.keys(order)[0]
                }
                else if (Object.keys(order).length == 2) {
                  if (order[Object.keys(order)[0]] < order[Object.keys(order)[1]] ) {
                    tdg1 = Object.keys(order)[0]
                    tdg2 = Object.keys(order)[1] 
                  }
                  else {
                    tdg1 = Object.keys(order)[1]
                    tdg2 = Object.keys(order)[0] 
                  }
                }
                else if (Object.keys(order).length == 3) {
                  
                  if (order[Object.keys(order)[0]] < order[Object.keys(order)[1]]) {

                    if (order[Object.keys(order)[1]] < order[Object.keys(order)[2]]) {
                      tdg1 = Object.keys(order)[0]
                      tdg2 = Object.keys(order)[1]
                      tdg3 = Object.keys(order)[2]
                    }
                    else {
                      if (order[Object.keys(order)[0]] < order[Object.keys(order)[2]]) {
                        tdg1 = Object.keys(order)[0]
                        tdg2 = Object.keys(order)[2]
                        tdg3 = Object.keys(order)[1]
                      }
                      else {
                        tdg1 = Object.keys(order)[2]
                        tdg2 = Object.keys(order)[0]
                        tdg3 = Object.keys(order)[1]
                      }
                    }
                  }
                  else { 
                      if (order[Object.keys(order)[0]] < order[Object.keys(order)[2]]) {
                      tdg1 = Object.keys(order)[1]
                      tdg2 = Object.keys(order)[0]
                      tdg3 = Object.keys(order)[2]
                    }
                    else {
                      if (order[Object.keys(order)[1]] < order[Object.keys(order)[2]]) {
                        tdg1 = Object.keys(order)[1]
                        tdg2 = Object.keys(order)[2]
                        tdg3 = Object.keys(order)[0]
                      }
                      else {
                        tdg1 = Object.keys(order)[2]
                        tdg2 = Object.keys(order)[1]
                        tdg3 = Object.keys(order)[0]
                      }
                    } 
                  }
                }
              }
              
              catch {
                Logger.log('(dot) transport info format issue:' +err+'\n')
              } 
            }

            // DOT GUIDE
            else if (headers[i] == 'DOT%20Emergency%20Guidelines') {
              try {
                for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Section[0].Information).length; j++) {
                  
                  try {
                    if (dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.match(/guide\s*\d+/i)[0].toLowerCase()) {
                      dot_guide += dataSiteCompound.Record.Section[0].Section[0].Section[0].Information[j].Value.StringWithMarkup[0].String.match(/guide\s*\d+/i)[0].toLowerCase()
                      break
                    }
                  }
                  catch {
                    Logger.log('dot guide bad format')
                  }

                }
              }
              catch (err) { 
                Logger.log('dot guide format issue:' +err+'\n')
              } 
              // Logger.log('dot guide: '+dot_guide)
            }

            // ERG HAZARD CLASSES
            else if (headers[i] == 'Hazards%20Identification') {
              for (let j = 0; j < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Information).length; j++) {
                try {        
                  if (dataSiteCompound.Record.Section[0].Section[0].Information[j].Name == 'ERG Hazard Classes') { 
                    for (let k = 0; k < Object.keys(dataSiteCompound.Record.Section[0].Section[0].Information[j].Value.StringWithMarkup).length; k++) {
                      erg_hc += dataSiteCompound.Record.Section[0].Section[0].Information[j].Value.StringWithMarkup[k].String+';'
                    }
                    break
                  }
                }
                catch (err) { 
                  Logger.log('erg_hc format issue:' +err+'\n')
                }
              }
              Logger.log('erg hazard: '+erg_hc)
            }
          } 
        }

        // HEADER NOT FOUND
        catch (err){
          Logger.log('Header not found: '+headers[i]+'\n'+err);          
          if (headers[i] == 'UN%20Classification') {
            tdg_flag = true;
          }
        }
      }

      // super entry      
      // pubchem cid
      if (cid) {
        range.setValue(cid);
      }
      // chemical desc.
      if (name) {
        range.offset(0,1).setValue(name);
      }
      // cas no.
      if (rn && !flag_rn) {
        range.offset(0,4).setValue(rn);
      }
      // synonyms
      if (syn) {
        range.offset(0,2).setValue(syn);
      }
      // chemical formula
      if (cf) {
        range.offset(0,9).setValue(cf);
      }
      // flash point
      if (fp) {
        range.offset(0,12).setValue(fp);
      }
      // density weight
      if (dw) {
        range.offset(0,13).setValue(dw);
      }
      // melting point
      if (mp) {
        range.offset(0,14).setValue(mp);
      }
      // boiling point
      if (bp) {
        range.offset(0,15).setValue(bp);
      }
      // molecular weight
      if (mw) {
        range.offset(0,16).setValue(mw);
      }
      // pH value (NOT USED)
      if (pH) {
        range.offset(0,17).setValue(pH);
      }
      // unit (NOT USED, could be set via physical state)
      if (unit) {
        range.offset(0,23).setValue(unit);
      }
      // physical state 
      if (ps) {
        range.offset(0,24).setValue(ps);
      }
      // canonical smiles
      if (cs) {
        range.offset(0,25).setValue(cs);
      }
      // special hazard
      if (sh) {
        range.offset(0,26).setValue(sh);
      }
      // compatibility category (USED, fill with reactive group and chemical classes)
      if (cc) {
        range.offset(0,27).setValue(cc);
      }
      // ERG Hazard Classes
      if (erg_hc) {
        range.offset(0,30).setValue(erg_hc);
      }
      // UN number
      if (un_no) {
        range.offset(0,31).setValue(un_no);
      }
      // UN Pack Group
      if (un_pg) {
        range.offset(0,32).setValue(un_pg);
      }
      // TDG primary
      if (tdg1) {
        range.offset(0,33).setValue(tdg1);
      }
      // TDG secondary
      if (tdg2) {
        range.offset(0,34).setValue(tdg2);
      }
      // TDG tertiary 
      if (tdg3) {
        range.offset(0,35).setValue(tdg3);
      }
      // storage requirements (NOT USED, for now)
      if (strReq) {
        range.offset(0,36).setValue(strReq);
      }
      // dot guide 
      if (dot_guide) {
        range.offset(0,37).setValue(dot_guide);
      }      
      // GHS classification
      if (ghs_class) {
        range.offset(0,38).setValue(ghs_class);
      }
      // GHS hazard statements
      if (ghs_hs) {
        range.offset(0,39).setValue(ghs_hs);
      }
      // GHS precautionary statement codes
      if (ghs_psc) {
        range.offset(0,40).setValue(ghs_psc);
      }
      // GHS pictogram(s)
      if (ghs_pic) {
        range.offset(0,41).setValue(ghs_pic);
      }
      // GHS signal word
      if (ghs_signal) {
        range.offset(0,42).setValue(ghs_signal);
      }
      // nfpa diamond
      if (nfpa_dia) {
        range.offset(0,43).setValue(nfpa_dia);
      }
      range.offset(0,44).setValue('y');
      range.offset(0,45).setValue('n');

      // sub entry(s) based on unique un_no's   
      var un_uniques = [] // specfic subsequent entry will have this un_no // this will be equal to the un_count
      var un_generics = [] // every subsequent entry will have these un_no's
      var un_temp = un_no.split(';')

      // Logger.log('un_no: '+un_no) 
      // Logger.log('un split: '+ un_temp+'\nlength: '+un_temp.length)

      for (let k = 0; k < un_no.split(';').length-1; k++) {
       
        // Logger.log('current: '+un_temp[k])
        // if (un_temp[k].match(/\((.*)\)/)) {
        //   Logger.log('regex: '+ un_temp[k].match(/\((.*)\)/)[1])
        // }
    
        if (un_temp[k].match(/\((.*)\)/) && un_temp[k].match(/\((.*)\)/)[1] && un_temp[k].match(/\((.*)\)/)[1].toLowerCase() !== name.toLowerCase() && !syn.toLowerCase().includes(un_temp[k].match(/\((.*)\)/)[1].toLowerCase())) { // check if un_no is a unique entry
          // Logger.log('unique')
          un_uniques.push(un_temp[k])
        }
        else {
          // Logger.log('generic')
          un_generics.push(un_temp[k])
        }
      }  
      // Logger.log('unique: '+un_uniques)
      // Logger.log('generic: '+un_generics)

      if (un_uniques.length > 1) { // must be greater than one or else we should only have one super entry that incl. this 
        SpreadsheetApp.getActiveSheet().insertRowsAfter(range.getRowIndex(),un_uniques.length)
        for (let i = 0; i < un_uniques.length; i++) {
          range = range.offset(1,0).activate()
          // pubchem cid
          if (cid) {
            range.setValue(cid);
          }
          // chemical desc.
          if (name) {
            range.offset(0,1).setValue(un_uniques[i].match(/\((.*)\)/)[1]);
          }
          // cas no.
          if (rn) {
            range.offset(0,4).setValue(rn);
          }
          // synonyms
          if (syn) {
            range.offset(0,2).setValue(syn);
          }
          // chemical formula
          if (cf) {
            range.offset(0,9).setValue(cf);
          }
          // flash point
          if (fp) {
            range.offset(0,12).setValue(fp);
          }
          // density weight
          if (dw) {
            range.offset(0,13).setValue(dw);
          }
          // melting point
          if (mp) {
            range.offset(0,14).setValue(mp);
          }
          // boiling point
          if (bp) {
            range.offset(0,15).setValue(bp);
          }
          // molecular weight
          if (mw) {
            range.offset(0,16).setValue(mw);
          }
          // pH value (NOT USED)
          if (pH) {
            range.offset(0,17).setValue(pH);
          }
          // unit (NOT USED, could be set via physical state)
          if (unit) {
            range.offset(0,23).setValue(unit);
          }
          // physical state // TODO: fix this based on name
          if (ps) {
            range.offset(0,24).setValue(ps);
          }
          // canonical smiles
          if (cs) {
            range.offset(0,25).setValue(cs);
          }
          // special hazard
          if (sh) {
            range.offset(0,26).setValue(sh);
          }
          // compatibility category (USED, fill with reactive group and chemical classes)
          if (cc) {
            range.offset(0,27).setValue(cc);
          }
          // ERG Hazard Classes
          if (erg_hc) {
            range.offset(0,30).setValue(erg_hc);
          }  
          // UN number
          if (un_no) {
            if (un_generics.length > 0) {  
              range.offset(0,31).setValue(un_uniques[i]+';'+un_generics.join(';')+';'); 
            }
            else {
              range.offset(0,31).setValue(un_uniques[i]+';');
            }
          }
          // UN Pack Group
          if (un_pg) {
            range.offset(0,32).setValue(un_pg);
          }
          // TDG primary
          if (tdg1) {
            range.offset(0,33).setValue(tdg1);
          }
          // TDG secondary
          if (tdg2) {
            range.offset(0,34).setValue(tdg2);
          }
          // TDG tertiary 
          if (tdg3) {
            range.offset(0,35).setValue(tdg3);
          }
          // storage requirements (NOT USED, for now)
          if (strReq) {
            range.offset(0,36).setValue(strReq);
          }
          // dot guide 
          if (dot_guide) {
            range.offset(0,37).setValue(dot_guide);
          }      
          // GHS classification
          if (ghs_class) {
            range.offset(0,38).setValue(ghs_class);
          }
          // GHS hazard statements
          if (ghs_hs) {
            range.offset(0,39).setValue(ghs_hs);
          }
          // GHS precautionary statement codes
          if (ghs_psc) {
            range.offset(0,40).setValue(ghs_psc);
          }
          // GHS pictogram(s)
          if (ghs_pic) {
            range.offset(0,41).setValue(ghs_pic);
          }
          // GHS signal word
          if (ghs_signal) {
            range.offset(0,42).setValue(ghs_signal);
          }
          // nfpa diamond
          if (nfpa_dia) {
            range.offset(0,43).setValue(nfpa_dia);
          }
          range.offset(0,44).setValue('y');
          range.offset(0,45).setValue('y');
        }
      }

      Logger.log('PubChem matched compound: '+ name + ' cas: ' + rn);
    }

    catch (err) {
      Logger.log(err);
      Logger.log('PubChem failed to match compound: '+ name + ' cas: ' + rn);
      range.offset(0,44).setValue('n');
    }

    if (range.getRowIndex() < SpreadsheetApp.getActiveSheet().getLastRow()) { // do not make new row if end of content
      range = range.offset(1,0).activate();
    }
  }
}
