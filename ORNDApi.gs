
function orndApi () {
  
  var postPayload = {
    "client_id": "HD9V1j6sI1aLieZ5",
    "client_secret": "iKeTMUaVXew1EwqQQSSRb2AC6vbehWQY",
    "grant_type": "client_credentials",
    "scope": "officernd.api.read"
  };

  var options = {
    "method" : "post",
    "payload" : postPayload
  };
  
  var pre_response = UrlFetchApp.fetch('https://identity.officernd.com/oauth/token', options);
  var pre_response_json = JSON.parse(pre_response.getContentText());
  var access_token = pre_response_json.access_token; // return the access token 

  //everything above here is to create the access token to access the ORND API(Authentication)
  //==============================================================================================================//

  //needed to write in this format to fulfill ORND requirement
   var headers = {
      "Authorization" : "Bearer " + access_token
  };
  
  //this will get datas from the ORND server
  var params = {
    "method": "get",
    "headers": headers
  };


  //import Teams JSON (https://developer.officernd.com/#teams)  
  var post_response_team = UrlFetchApp.fetch('https://app.officernd.com/api/v1/organizations/worq/teams', params);
  var team_json = JSON.parse(post_response_team);
  var data_team = team_json;

  //import Membership JSON(https://developer.officernd.com/#memberships)
  var post_response_memberships = UrlFetchApp.fetch('https://app.officernd.com/api/v1/organizations/worq/memberships', params);
  var memberships_json = JSON.parse(post_response_memberships);
  var data_memberships = memberships_json;

   //import Members JSON(https://developer.officernd.com/#memberships)
  var post_response_members = UrlFetchApp.fetch('https://app.officernd.com/api/v1/organizations/worq/members', params);
  var members_json = JSON.parse(post_response_members);
  var data_members = members_json;

  //import Plans JSON(https://developer.officernd.com/#memberships)
  var post_response_plans = UrlFetchApp.fetch('https://app.officernd.com/api/v1/organizations/worq/plans', params);
  var plans_json = JSON.parse(post_response_plans);
  var data_plans = plans_json;

  // the API datas retrieval ends here
  //===============================================================================================================//

  var app= SpreadsheetApp.openById('1Np3UhwCmN9hqWHGrhrFOIxteeFaDkU55q0lWurSE_EQ');
  var ss = app.getSheetByName('Upload');
  
  //Making a new team array 
  const newData_team = data_team.map(({
  name: pic, //changing name-> pic
  _id: idTeam, //changing _id-> idTeam
  email: emailTeam, //changing email-> emailTeam
  ...rest
  }) => ({
  pic,
  idTeam,
  emailTeam,
  ...rest
  }));

  //Making a new member array 
  const newData_member = data_members.map(({
  team: idTeam, //changing team-> idTeam
  name: memberName, //changing name-> memberName
  _id: idMember, //changing _id-> idMember
  ...rest
  }) => ({
  idTeam,
  memberName,
  idMember,
  ...rest
  }));

  //Making a new plans array 
  const newData_plans = data_plans.map(({
  _id: idPlan, //changing _id-> idPlan
  name: planName, //changing name-> planName
  ...rest
  }) => ({
  idPlan,
  planName,
  ...rest
  }));

  //Making a new membership array 
  const newData_membership = data_memberships.map(({
  team: idTeam, //changing team-> idTeam
  price: priceMembership, //changing price-> priceMembership
  discountedPrice: discountedPriceMembership, //changing discountedPrice -> discountedPriceMembership
  plan: idPlan, //changing plan-> idPlan
  member: idMember, //changing member-> idMember
  ...rest
  }) => ({
  idTeam,
  priceMembership,
  discountedPriceMembership,
  idPlan,
  idMember,
  ...rest
  }));

  //merging every data into 1 array accodirng to idTeam and idPlan
  const mergeByIdTeam = (a1, a2) =>       
    a1.map(itm => ({
        ...a2.find((item) => (item.idTeam === itm.idTeam) && item),
        ...itm
    }));

  const mergeByIdPlan = (a1, a2) =>
    a1.map(itm => ({
        ...a2.find((item) => (item.idPlan === itm.idPlan) && item),
        ...itm
    }));  

  //const testtest = newData_membership.filter(x => x.calculatedStatus === 'active');

  let testArr = mergeByIdTeam(newData_membership, newData_member);
  let testArr2 = mergeByIdTeam(testArr, newData_team);
  let mergedArr = mergeByIdPlan(testArr2, newData_plans); //The final Array
  //end of merging

var today = new Date();

for (var x = 0 ; x < mergedArr.length ; x++) { //This for loop is only for formatting the table

    mergedArr.map(function(item){
      if(item.office === "565748274a955c790d808c77"){
      item.office = "WORQ Subang"
      }

      if(item.office === "5dac63c998e930010a595016"){
      item.office = "WORQ Gateway"
      }

      if(item.office === "5db8fb7e35798d0010950a77"){
      item.office = "WORQ TTDI"
      }

      if(item.office === "5db8fb9798549f0010df15f3"){
      item.office = "WORQ Surian"
      }

      return item;
    });

    var tempDate = new Date (mergedArr[x].endDate);

    var diffInMs = tempDate.getTime() - today.getTime();

    var diff = Math.round(diffInMs / (1000 * 3600 * 24));

    if(mergedArr[x].endDate != null){

      if( diff <= 30){

        if (tempDate < today){
          mergedArr[x].calculatedStatus = "Expired";
        } 
        else {
          mergedArr[x].calculatedStatus = "Expiring";
        }
      }

      if( diff > 30){
      mergedArr[x].calculatedStatus = "Not Expiring";
      } 

    } 
    else {
      mergedArr[x].calculatedStatus = "Not Expiring";
    } 

}


//Methods for reading the array and writing it to the sheet
var headings = ['pic', 'memberName',  'startDate', 'endDate', 'category',	'planName',	'priceMembership',	'discountedPriceMembership',	'office',	'calculatedStatus',	'emailTeam'];
var outputRows = [];

// Loop through 
mergedArr.forEach(function(i) {
  // Add a new row to the output mapping each header to the corresponding value.
  outputRows.push(headings.map(function(heading) {
    return i[heading] || '';
  }));
});

// Write to sheets
if (outputRows.length) {
  // Add the headings - delete this next line if headings not required
  outputRows.unshift(headings);
  ss.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
}
// end of reading and writing to sheet


//Changing the header 
var finalHeader = ['Team','Member', 'Start Date', 'End Date', 'Plan',	'Name',	'Price',	'Discounted Price',	'Location',	'Status',	'Email'];
ss.getRange(1,1,1,finalHeader.length).setValues([finalHeader]);

}





