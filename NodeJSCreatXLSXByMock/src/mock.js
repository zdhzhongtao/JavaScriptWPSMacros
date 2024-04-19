// 引入mockjs
const Mock = require("mockjs");
const XLSX = require('xlsx');
// 获取 mock.Random 对象
const Random = Mock.Random;
// 转换函数
function convertTo2DArray(objArray) {
  var keys = Object.keys(objArray[0]);
  var result = [keys]; // 将键作为第一行

  for (var i = 0; i < objArray.length; i++) {
      var values = [];
      for (var j = 0; j < keys.length; j++) {
          values.push(objArray[i][keys[j]]);
      }
      result.push(values);
  }

  return result;
}


// mock新闻数据，包括新闻标题title、内容content、创建时间createdTime
const produceRandomData = function () {
  let list = [];
  for (let i = 0; i < 20; i++) {
    let randomDataObject = {
        ProductID: Mock.mock('@id'), //
        //Mock.Random.pick(['ProductName1', 'ProductName2', 'ProductName3', 'ProductName4', 'ProductName5', 'ProductName6', 'ProductName7']), //
        ProductName: Mock.mock(()=>{return `Product_${i}`}),
        ProductGroup: Mock.Random.pick(['ProductGroup1', 'ProductGroup2', 'ProductGroup3']), //
        StandardCost: Mock.Random.pick(['StandardCost1', 'StandardCost2', 'StandardCost3', 'StandardCost4']), //
    };
    list.push(randomDataObject);
  }
  return list;
};
const produceRandomDataList = produceRandomData();
const produceIDRandomDataList = produceRandomDataList.map(obj => obj.ProductID);


// 生成用户名表
const customerRandomData = function () {
  let list = [];
  for (let i = 0; i < 120; i++) {
    let randomDataObject = {
      Customer: Mock.mock('@id'), //
      CustomerName: Mock.Random.cname(), //
    };
    list.push(randomDataObject);
  }
  return list;
};
const customerRandomDataList = customerRandomData();
const customerIDRandomDataList = customerRandomDataList.map(obj => obj.Customer);

// 生成Label表
const labelRandomData = function () {
  let list = [];
  for (let i = 0; i < 20; i++) {
    let randomDataObject = {
      LabelID: Mock.mock('@id'), //
      Label: Mock.Random.cname(), //
    };
    list.push(randomDataObject);
  }
  return list;
};
const labelRandomDataList = labelRandomData();


// 生成用户名表
const LabelMappingRandomData = function () {
  let list = [];
  for (let i = 0; i < 20; i++) {
    let randomDataObject = {
      Customer: Mock.Random.pick(customerIDRandomDataList), // 只从Customer表中获取存在的数据
      LabelID: Mock.Random.pick(labelRandomDataList.map(obj => obj.LabelID)), //
      TimeStamp: Mock.Random.datetime("yyyy-MM-dd HH:mm:ss"), //Mock.Random.datetime("yyyy-MM-dd HH:mm:ss")
    };
    list.push(randomDataObject);
  }
  return list;
};


// 生成Channel表
const ChannelRandomData = function () {
  let list = [];
  for (let i = 0; i < 20; i++) {
    let randomDataObject = {
      ChannelID: Mock.mock('@id'), //
      Name: Mock.Random.cname(), //
    };
    list.push(randomDataObject);
  }
  return list;
};
const channelRandomDataList = ChannelRandomData();
const channelIDRandomDataList = channelRandomDataList.map(obj => obj.ChannelID);

// 生成Sales订单表  ChannelID	OrderDate	Customer	Country	Province RelatedCoupon	Status
const SalesMappingRandomData = function () {
  let list = [];
  for (let i = 0; i < 200; i++) {
    let randomDataObject = {
      OrderID:  Mock.mock('@id'), //
      ChannelID: Mock.Random.pick(channelIDRandomDataList), //
      OrderDate: Mock.Random.datetime("yyyy-MM-dd HH:mm:ss"), //Mock.Random.datetime("yyyy-MM-dd HH:mm:ss")
      Customer: Mock.Random.pick(customerIDRandomDataList), 
      Country: Mock.Random.pick(["成都市","北京市","上海市","天津市"]), 
      Province: Mock.Random.pick(["四川省","北京市","上海市","天津市"]), 
      RelatedCoupon: Mock.Random.pick(customerIDRandomDataList), 
      RelatedCoupon: Mock.Random.pick(["新订单","已支付","已送达","已退货"]), 
    };
    list.push(randomDataObject);
  }
  return list;
};
const salesMappingRandomDataList = SalesMappingRandomData()
const salesMappingIDRandomData = salesMappingRandomDataList.map(obj => obj.OrderID);



// 生成SalesLine表  OrderID	Product	Qty	Price	Discount	Amount

const SalesLineRandomData = function () {
  let list = [];
  for (let i = 0; i < 200; i++) {
    let randomDataObject = {
      OrderID: Mock.Random.pick(salesMappingIDRandomData), //
      Product: Mock.Random.pick(produceIDRandomDataList), //
      Qty: Mock.Random.integer(1, 5), //
      Price:Mock.Random.float(10, 100, 0, 2), //
      Discount: Mock.Random.float(10, 100, 0, 2), //
      Amount: Mock.Random.float(10, 100, 0, 2), //
    };
    list.push(randomDataObject);
  }
  return list;
};
const salesLineRandomDataList = SalesLineRandomData();


//ReturnID	OrderID	ReturnDate	ReturnReason	Comments

const ReturnOrderRandomData = function () {
  let list = [];
  for (let i = 0; i < 200; i++) {
    let randomDataObject = {
      ReturnID: Mock.mock('@id'), //
      OrderID: Mock.Random.pick(salesMappingIDRandomData), //
      ReturnDate: Mock.Random.datetime("yyyy-MM-dd HH:mm:ss"), //
      ReturnReason:Mock.Random.pick(["商返","七天无理由","质量问题","召回"]), //
      Comments: "", //
    };
    list.push(randomDataObject);
  }
  return list;
};
const returnOrderRandomDataList = ReturnOrderRandomData();
const returnOrderIDRandomDataList = returnOrderRandomDataList.map(obj => obj.ReturnID);

//ReturnID	Product	Qty	Amount

const ReturnOrderLineRandomData = function () {
  let list = [];
  for (let i = 0; i < 200; i++) {
    let randomDataObject = {
      ReturnID: Mock.Random.pick(returnOrderIDRandomDataList), //
      Product: Mock.Random.pick(salesMappingIDRandomData), //
      Qty: Mock.Random.integer(1, 5), //
      Amount:Mock.Random.integer(10, 50), //
    };
    list.push(randomDataObject);
  }
  return list;
};


// 创建一个工作簿
var wb = XLSX.utils.book_new();

// 将你的数组转换为工作表
var produceRandomDataListWS = XLSX.utils.aoa_to_sheet(convertTo2DArray(produceRandomDataList));
// 将工作表添加到工作簿
XLSX.utils.book_append_sheet(wb, produceRandomDataListWS, 'Product');

var customerRandomDataListWS = XLSX.utils.aoa_to_sheet(convertTo2DArray(customerRandomDataList));
XLSX.utils.book_append_sheet(wb, customerRandomDataListWS, 'Customer');

var labelRandomDataListWS = XLSX.utils.aoa_to_sheet(convertTo2DArray(labelRandomDataList));
XLSX.utils.book_append_sheet(wb, labelRandomDataListWS, 'Label');

var labelMappingRandomDataListWS = XLSX.utils.aoa_to_sheet(convertTo2DArray(LabelMappingRandomData()));
XLSX.utils.book_append_sheet(wb, labelMappingRandomDataListWS, 'labelMapping');

var salesMappingRandomDataListWS = XLSX.utils.aoa_to_sheet(convertTo2DArray(salesMappingRandomDataList));
XLSX.utils.book_append_sheet(wb, salesMappingRandomDataListWS, 'SalesMapping');

var channelRandomDataListWS = XLSX.utils.aoa_to_sheet(convertTo2DArray(channelRandomDataList));
XLSX.utils.book_append_sheet(wb, channelRandomDataListWS, 'Channel');

var salesLineRandomDataListWS = XLSX.utils.aoa_to_sheet(convertTo2DArray(salesLineRandomDataList));
XLSX.utils.book_append_sheet(wb, salesLineRandomDataListWS, 'SalesLine');

var returnOrderRandomDataListWS = XLSX.utils.aoa_to_sheet(convertTo2DArray(returnOrderRandomDataList));
XLSX.utils.book_append_sheet(wb, returnOrderRandomDataListWS, 'ReturnOrder');

var returnOrderLineRandomDataListWS = XLSX.utils.aoa_to_sheet(convertTo2DArray(ReturnOrderLineRandomData()));
XLSX.utils.book_append_sheet(wb, returnOrderLineRandomDataListWS, 'ReturnOrderLine');

// 将工作簿写入文件
XLSX.writeFile(wb, 'output.xlsx');