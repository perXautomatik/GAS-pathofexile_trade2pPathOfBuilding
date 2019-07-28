
function JsonToItemz() {
    
  const spr = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spr.getActiveSheet();
  const column = sheet.getRange('B2:K2');
  const values = column.getValues(); // get all data in one call
  var Jsons = [];
  
  values[0].forEach(function (e) {
    var temp = e;
    var jzons = JSON.parse(temp);
    Jsons.unshift(jzons);   
  });
  
  var Item = function(e){
      
    this.HDivider = "		<Item>";
   
    this.Rarity = "			Rarity: " + 
      ((("undefined" !== typeof(e.item.flavourText)) 
       ? "Unique" :
       (("undefined" !== typeof(e.item['name'])) 
       ? "Rare" : "Magic"))).toUpperCase();
    
    if(this.Rarity !== "MAGIC" AND this.Rarity !== "NORMAL")
    {
          this.UniqueID = "Unique ID: " + e.id;
          this.itemName = e.item['name'] || e.item['typeLine'];
    }
        
    this.ItemType = e.item['typeLine'] || e.item.category[Object.keys(e.item.category)[0]][0] || "jewel";
    this.ItemType = (this.ItemType === "abyss" ? "abyssal jewel" : this.ItemType);  
    

    this.ItemLvl = "Item Level: " + e.item.ilvl;
      
    if("undefined" !== typeof(e.item.properties)){
      if(e.item.properties[0].name === "Quality"){
        this.Quality = "Quality: " + e.item.properties[0].values[0];
      }
    }
    
    if("undefined" !== typeof(e.item.sockets)){
      var socks="";
      e.item.sockets.forEach(function (it) {socks += it.sColour;});
      this.Sockets = "Sockets: " + socks.split('').join('-');
    }
    
      this.LevelReq = "LevelReq: " + parseInt((("undefined" === typeof( e.item["requirements"])) ? 0 : e.item.requirements[0].values[0]));
    if("undefined" !== typeof(e.item.implicitMods)){   
      this.ImplicitsNr = "Implicits: " + (("undefined" === typeof( e.item.implicitMods)) ? 0 : e.item.implicitMods.length);
      this.implicit = e.item.implicitMods;
    }
      this.modifier = e.item.explicitMods;
    this.StashPosition = encodeURI(JSON.stringify(e.listing.stash)).replace(/\W|22|7D|7B/g, ' ').replace(/\s{2,}/g, ' ');
      this.BDivider = "		</Item>";
 };
  
  var itemz = [];
  Jsons.forEach(function (e) {
    e.result.forEach(function (g) {
      itemz.unshift(new Item(g));});});
  
  var matrix = [];
  var i = 0;
  itemz.reverse();
  itemz.forEach(function (it) {
    
    for (var p in it) {if (it.hasOwnProperty(p)) 
    {
      matrix[i] = [];

      if (Array.isArray(it[p]))
      {
        it[p].forEach(function (g) {
          var tempx = [];tempx.unshift(g);matrix[i] = tempx;i++;
        });
      }
      else{matrix[i++][0] = it[p];}
      
    }
    }}); 
  
   var ressult = sheet.getRange(1, 1, matrix.length, 1);
  ressult.setValues(matrix); 
}


