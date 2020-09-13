class HolidaysWP {
  constructor() {
    this.holidaysList;
    this.holidaysBodyId = "dprtBody";
    this.listTitle = "Holidays";
    this.endPoint = `/_api/lists/getbytitle('${this.listTitle}')/items`;
  }

  async getData() {
    //************* ВАШ КОД ТУТ *************
    const query = `${_spPageContextInfo.webAbsoluteUrl}${this.endPoint}?$select=Id,Title,Category,IsNonWorkingDay,V4HolidayDate,Worker/Title,Worker/Id,Places/Id,Places/Title&$expand=Places,Worker`;
    const request = await fetch(query, {
      headers: {
        accept: "application/json;odata=verbose"
      }
    });
    const result = await request.json();
    this.renderHTML(result);
    //***************************************
  }

  async addItem() {
    const requestDigest = await this.getRequestDigest(_spPageContextInfo.webAbsoluteUrl);
    const listItemType = await this.getListItemType(_spPageContextInfo.webAbsoluteUrl, this.listTitle);
    const obj = {__metadata: {
      type: "Collection(Edm.Int32)"
    }};
    obj.results = [1, 2, 3];
    const newItem = {
      Title: "Default",
      Category: "Summer",
      IsNonWorkingDay: true,
      WorkerId: 12,
      V4HolidayDate: (new Date()).toISOString(),
      PlacesId: obj
    }
    const objType = {
      __metadata: {
        type: listItemType.d.ListItemEntityTypeFullName,
      },
    };
    const objData = JSON.stringify(Object.assign(objType, newItem));
    return $.ajax({
      url: `${_spPageContextInfo.webAbsoluteUrl}${this.endPoint}`,
      type: "POST",
      data: objData,
      headers: {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest":
          requestDigest.d.GetContextWebInformation.FormDigestValue,
        "X-HTTP-Method": "POST",
      },
    });

  }

  async updateItem() {
    const query = `${_spPageContextInfo.webAbsoluteUrl}${this.endPoint}(6)`;
    const requestDigest = await this.getRequestDigest(_spPageContextInfo.webAbsoluteUrl);
    const listItemType = await this.getListItemType(_spPageContextInfo.webAbsoluteUrl, this.listTitle);

    const changes = {
      Title: "Eeeeee",
    };
    const objType = {
      __metadata: {
        type: listItemType.d.ListItemEntityTypeFullName,
      },
    };
    const objData = JSON.stringify(Object.assign(objType, changes));
    return $.ajax({
      url: query,
      type: "POST",
      data: objData,
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest":
          requestDigest.d.GetContextWebInformation.FormDigestValue,
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE",
      }
    });
  }


  getRequestDigest(url) {
    return $.ajax({
      url: `${url}/_api/contextinfo`,
      method: "POST",
      headers: {
        Accept: "application/json; odata=verbose",
      },
    });

  }

  getListItemType(url, listTitle) {
    return $.ajax({
      url: `${url}/_api/Web/Lists/getbytitle('${listTitle}')/ListItemEntityTypeFullName`,
      method: "GET",
      contentType: "application/json;odata=verbose",
      headers: {
        Accept: "application/json;odata=verbose",
      }
    });
  }

  renderHTML(result) {
    try {
      const holidaysList = result.d.results;
      let holidaysItems = "";
      holidaysList.map((item) => {
        const holidayItem = new HolidayItem(item);
        holidaysItems += holidayItem.getHTML();
      });
      document.getElementById(
        this.holidaysBodyId
      ).innerHTML = holidaysItems;

    } catch (error) {
      this.webpartNoDataContainer("dprtWrapper", "No data provided");
    }
  }

  webpartNoDataContainer(domName, errorMessage) {
    let errorHTML = `<div class="noDataContainer">
        <div id="warning-block">
            <img src="../SiteAssets/anshy-site-assets/noData.svg" alt="error"/>
        </div>
      </div>`;
    document.getElementById(domName).innerHTML = errorHTML;
    console.error(`Error in ${domName} -> ${errorMessage}`);
  }
}

class HolidayItem {
  constructor(holidayItem) {
    this.id = holidayItem.Id;
    this.title = holidayItem.Title;
    this.category = holidayItem.Category;
    this.working = holidayItem.IsNonWorkingDay;
    this.dateTime = new Date(holidayItem.V4HolidayDate).toDateString();
    this.worker = holidayItem.Worker.Title;
    this.places = "";
    for(let place of holidayItem.Places.results) {
      this.places += `<p class="dprtCard__text"><i>${place.Title}</i></p>`
    }
  }

  getHTML() {
    let holidayItemTemplate = `
      <div class="dprtCard dprtCard_hover" data-id=${this.id}>
        <h3 class="dprtCard__name">${this.title}</h3>
        <p class="dprtCard__text">Category: ${this.category}</p>
        <p class="dprtCard__text">Non-working: ${this.working}</p>
        <p class="dprtCard__text">Date: ${this.dateTime}</p>
        <p class="dprtCard__text">Worker: ${this.worker}</p>
        <p class="dprtCard__text"><b>PLACES</b></p>
        ${this.places}
      </div>`;
    return holidayItemTemplate;
  }
}

SP.SOD.executeFunc("sp.js", "SP.ClientContext", async function () {
  const dprtWP = new HolidaysWP();
  dprtWP.getData();
  document.querySelector(".add").addEventListener("click", async () => {
    await dprtWP.addItem();
    await dprtWP.getData();
  })
});
