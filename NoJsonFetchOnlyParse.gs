
function JsonToItemz() {

    const spr = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spr.getActiveSheet();
    const column = sheet.getRange('B2:K2');
    const values = column.getValues(); // get all data in one call
    var Jsons = [];

    values[0].forEach(function (e) {
      try{
        Jsons.unshift(JSON.parse(e));
      }
      catch (e)
      {
        console.log(e);
      }
    });

    var Item = function (jz) {
     
        this["HDivider"] = "<Item>";

        var temp = jz.item['name'] || jz.item['typeLine'];

        var s = ("undefined" !== typeof (jz.item["flavourText"]) ? "UNIQUE" : (temp === jz.item['typeLine'] ? "MAGIC" : "RARE"));
        this["Rarity"] = "Rarity: " + s;

        if (s !== "MAGIC" && s !== "NORMAL") {
            this["itemName"] = temp;
        }

        this.ItemType = jz.item['typeLine'] || jz.item.category[Object.keys(jz.item.category)[0]][0] || "jewel";
        this.ItemType = (this.ItemType === "abyss" ? "abyssal jewel" : this.ItemType);

        if (s !== "MAGIC" && s !== "NORMAL") {
            this["UniqueID"] = "Unique ID: " + jz.id;
        }

        this["ItemLvl"] = "Item Level: " + jz.item.ilvl;

        if ("undefined" !== typeof (jz.item.properties) && jz.item.properties[0].name === "Quality") {
            this.Quality = "Quality: " + jz.item.properties[0].values[0];
        }

        if ("undefined" !== typeof (jz.item.sockets)) {
            var socks = "";
            jz.item.sockets.forEach(function (it) {
                socks += it.sColour;
            });
            this["Sockets"] = "Sockets: " + socks.split('').join('-');
        }

        this.LevelReq = "LevelReq: " + parseInt((("undefined" === typeof (jz.item["requirements"])) ? 0 : jz.item.requirements[0].values[0]));

        if ("undefined" !== typeof (jz.item.implicitMods)) {
            this.ImplicitsNr = "Implicits: " + (("undefined" === typeof (jz.item.implicitMods)) ? 0 : jz.item.implicitMods.length);
            this.implicit = jz.item.implicitMods;
        }
        this.modifier = jz.item.explicitMods;
        this.StashPosition = encodeURI(JSON.stringify(jz.listing.stash)).replace(/\W|22|7D|7B/g, ' ').replace(/\s{2,}/g, ' ');
        this.BDivider = "</Item>";
    };

    var itemz = [];
    Jsons.forEach(function (e) {
        e.result.forEach(function (g) {
            itemz.unshift(new Item(g));
        });
    });

    var matrix = [];
    var i = 0;
    itemz.reverse();
    itemz.forEach(function (it) {

        for (var p in it) {
            if (it.hasOwnProperty(p)) {
                matrix[i] = [];

                if (Array.isArray(it[p])) {
                    it[p].forEach(function (g) {
                        var tempx = [];
                        tempx.unshift(g);
                        matrix[i] = tempx;
                        i++;
                    });
                } else {
                    matrix[i++][0] = it[p];
                }

            }
        }
    });

  
    
  matrix.unshift(['<SharedItems>']);
  matrix.unshift(['</Accounts>']);
  matrix.unshift(['<Account accountName="PerAutomatik" sessionID="c7b33a33efc73853c63228a603ac3c0d"/>']);
  matrix.unshift(['<Accounts lastAccountName="PerAutomatik" lastRealm="PC">']);
  matrix.unshift(['</Mode>']);
  matrix.unshift(['<Arg string="asda"/>']);
  matrix.unshift(['<Arg string=\"E:\\Users\\Dator\\Documents/Path of Building/Builds/asda.xml"/>']);
  matrix.unshift(['<Mode mode="BUILD">']);
  matrix.unshift(['<PathOfBuilding>']);
  matrix.unshift(['<?xml version="1.0" encoding="UTF-8"?>']);
  
	matrix[matrix.length] =['</SharedItems>'];
	matrix[matrix.length] =['<Misc showThousandsCalcs="false" showThousandsSidebar="false" nodePowerTheme="RED/BLUE" buildSortMode="NAME"/>'];
  matrix[matrix.length] =['</PathOfBuilding>'];
  
    var ressult = sheet.getRange(1, 1, matrix.length, 1);
    ressult.setValues(matrix);
}


