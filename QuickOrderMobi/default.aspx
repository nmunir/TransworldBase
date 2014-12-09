<%@ Page Language="VB" ValidateRequest="false" %>

<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Web.UI" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Transworld Ordering System</title>
    <link rel="stylesheet" href="http://code.jquery.com/mobile/1.2.0/jquery.mobile-1.2.0.min.css" />
    <script type="text/javascript" src="http://code.jquery.com/jquery-1.8.2.min.js"></script>
    <script type="text/javascript" src="http://code.jquery.com/mobile/1.2.0/jquery.mobile-1.2.0.min.js"></script>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <style type="text/css">
        .Error
        {
            color: Red;
        }
        
        .Notification
        {
            color: Blue;
        }
        
        input.ui-focus, textarea.ui-focus
        {
            outline: none;
            -webkit-box-shadow: none;
        }
        em
        {
            color: red;
            font-weight: bold;
            padding-right: .25em;
        }
        .clear
        {
            clear: both;
        }
        .hide
        {
            display: none;
        }
        
        img
        {
            margin-top: 15px;
            margin-left: 5px;
            margin-bottom: 5px;
        }
        
        .Basket
        {
            margin-left: 15px;
            margin-top: -10px;
        }
    </style>
    <script type="text/javascript">


        

        function ClearFormFields() {

            $('form :input').val("");

        }

        function ClearSession() {

            sessionStorage.UserID = null;

        }

        function ClearBasket() {

            sessionStorage.Basket = null;

        }


        function getPageName() {

            urlStr = window.location.href;
            var pageName = "";
            if (urlStr.indexOf("#") > -1) {
                var param = urlStr.split("#");
                pageName = param[1];
            }

            return pageName;

        }


        function SetFocus() {
            
            if (localStorage.UserID != "null" && localStorage.UserID != "") {                

                $("#txtUserID").val(localStorage.UserID);
                $("#txtPassword").focus();

            }
            else 
            {

                $("#txtUserID").focus();
            }

            

        }

        $(function () {


            SetFocus();

            if (sessionStorage.UserID != "null") {

                var pageName = getPageName();

                if (pageName != null && pageName != "") {

                    if (pageName.toLowerCase() == "pgproducts") {
                        PageMethods.GetProducts(sessionStorage.UserID, OnSuccessGetProducts, OnErrorGetProducts);
                        $.mobile.changePage("#pgProducts");
                    }
                    else if (pageName.toLowerCase() == "pgproductdetail") {
                        $.mobile.changePage("#pgProductDetail");
                        PageMethods.GetProdInfoByID(sessionStorage.CurrentProductKeyInBasket, OnSuccessGetProdInfoByID, OnErrorGetProdInfoByID);

                    }
                    else if (pageName.toLowerCase() == "pgbasket") {

                        $.mobile.changePage("#pgBasket");
                        var itemStrings = sessionStorage.Basket;
                        var sUserID = sessionStorage.UserID;
                        PageMethods.CreateBasketFromString(itemStrings, sUserID, OnSuccessCreateBasketFromString, OnErrorCreateBasketFromString);

                    }

                    else if (pageName.toLowerCase() == "pgaddress") {

                        $.mobile.changePage("#pgaddress");
                        PageMethods.GetCountries(OnSuccessGetCountries, OnErrorGetCountries);
                    }


                }

            }
            else {
                $.mobile.changePage("#pgLogin");
            }


            $("#productsInBasket").on('click', function () {

                alert("hi");
                var nLogisticProductKey = $(this).attr("ID");
                var quantity = $(this).children("span").attr("html");
                alert(quantity);
                PageMethods.GetProdInfoByID(nLogisticProductKey, OnSuccessGetProdInfoByID, OnErrorGetProdInfoByID);
                $.mobile.changePage("#pgProductDetail");

            });

            $("#btnLogin").click(function () {

                var sUserID = $("#txtUserID").val();
                var sPassword = $("#txtPassword").val();
                PageMethods.VerifyUserCredentials(sUserID, sPassword, OnSuccessVerifyPassword, OnErrorVerifyPassword);
                return false;

            });

            $("#productList").on('click', 'li', function () {

                var nLogisticProductKey = $('img', this).attr("id");
                sessionStorage.CurrentProductKeyInBasket = nLogisticProductKey;
                $.mobile.changePage("#pgProductDetail");
                PageMethods.GetProdInfoByID(nLogisticProductKey, OnSuccessGetProdInfoByID, OnErrorGetProdInfoByID);
                ClearFormFields();

            });

            $("#btnBackProductDetail").click(function () {

                if (sessionStorage.UserID != null) {
                    PageMethods.GetProducts(sessionStorage.UserID, OnSuccessGetProducts, OnErrorGetProducts);
                }
                else {

                    $.mobile.changePage("#pgLogin");
                }

                return false;

            });

            $("#btnBackBasket").click(function () {

                if (sessionStorage.UserID != null) {
                    PageMethods.GetProducts(sessionStorage.UserID, OnSuccessGetProducts, OnErrorGetProducts);

                }
                else {

                    $.mobile.changePage("#pgLogin");

                }
                return false;

            });


            $("#imgProduct").click(function () {

                var src = $(this).attr("src").replace("thumbs", "jpgs");
                $("#imgProductPopup").attr("src", src);
                $("#divPopupImage").popup("open");

            });

            $("#btnConfirmOrder").click(function () {

                if (sessionStorage.UserID != null && sessionStorage.Basket != null) {

                    var companyName = $("#txtCompanyName").val();

                    var ctcName = $("#txtContactName").val();

                    var address1 = $("#txtAddress1").val();

                    var address2 = $("#txtAddress2").val();

                    var town = $("#txtTown").val();

                    var postCode = $("#txtPostCode").val();

                    var custRef1 = $("#txtCustRef1").val();

                    var specialInst = $("#txtSpecialInstructions").val();

                    var selectCountry = document.getElementById('selectCountry');

                    var countryKey = selectCountry.options[selectCountry.selectedIndex].value;

                    if ($("#txtCompanyName").val() == '') {

                        $("#lblCompanyNameError").html("please enter company name");
                        $("#lblCompanyNameError").addClass("Error");

                    }
                    else {

                        $("#lblCompanyNameError").html("");
                        $("#lblCompanyNameError").removeClass("Error");

                    }

                    if ($("#txtContactName").val() == '') {

                        $("#lblContactNameError").html("please enter contact name");
                        $("#lblContactNameError").addClass("Error");

                    }
                    else {

                        $("#lblContactNameError").html("");
                        $("#lblContactNameError").removeClass("Error");
                    }

                    if ($("#txtAddress1").val() == '') {

                        $("#lblAddress1Error").html("please enter address1");
                        $("#lblAddress1Error").addClass("Error");

                    }
                    else {

                        $("#lblAddress1Error").html("");
                        $("#lblAddress1Error").removeClass("Error");

                    }

                    if ($("#txtAddress2").val() == '') {

                        $("#lblAddress2Error").html("please enter address2");
                        $("#lblAddress2Error").addClass("Error");

                    }
                    else {

                        $("#lblAddress2Error").html("");
                        $("#lblAddress2Error").removeClass("Error");

                    }

                    if ($("#txtTown").val() == '') {

                        $("#lblTownError").html("please enter town");
                        $("#lblTownError").addClass("Error");

                    }
                    else {

                        $("#lblTownError").html("");
                        $("#lblTownError").removeClass("Error");

                    }

                    if ($("#txtPostCode").val() == '') {

                        $("#lblPostCodeError").html("please enter post code");
                        $("#lblPostCodeError").addClass("Error");

                    }
                    else {

                        $("#lblPostCodeError").html("");
                        $("#lblPostCodeError").removeClass("Error");

                    }

                    if ($("#txtCustRef1").val() == '') {

                        $("#lblCustRef1Error").html("please enter customer reference");
                        $("#lblCustRef1Error").addClass("Error");

                    }
                    else {

                        $("#lblCustRef1Error").html("");
                        $("#lblCustRef1Error").removeClass("Error");

                    }

                    if ($("#txtSpecialInstructions").val() == '') {

                        $("#lblSpecialInstError").html("please enter special Instructions");
                        $("#lblSpecialInstError").addClass("Error");

                    }
                    else {

                        $("#lblSpecialInstError").html("");
                        $("#lblSpecialInstError").removeClass("Error");

                    }

                    if ($("#selectCountry").val() == '') {

                        $("#lblSpecialInst").html("please select country");
                        $("#lblSpecialInst").addClass("Error");

                    }
                    else {

                        $("#lblSpecialInst").html("");
                        $("#lblSpecialInst").removeClass("Error");

                    }

                    if (countryKey == 0) {

                        $("#lblSelectCountryError").html("please select country");
                        $("#lblSelectCountryError").addClass("Error");

                    }

                    else {

                        $("#lblSelectCountryError").html("");
                        $("#lblSelectCountryError").removeClass("Error");

                    }

                    if (companyName != '' && ctcName != '' && address1 != '' && address2 != '' && town != '' && postCode != '' && custRef1 != '' && specialInst != '' && countryKey > 0) {

                        PageMethods.nSubmitConsignment(sessionStorage.UserID, companyName, ctcName, address1, address2, town, postCode, custRef1, specialInst, countryKey, sessionStorage.Basket, OnSuccessSubmitConsignment, OnErrorSubmitConsignment);

                    }

                    return false;

                }
                else {

                    $.mobile.changePage("#pgLogin");
                    ClearBasket();
                    ClearFormFields();
                    ClearSession();

                }

            });

            $("#btnShowBasket").click(function () {

                if (sessionStorage.UserID != null) {
                    $.mobile.changePage("#pgBasket");
                    if (sessionStorage.Basket != null && sessionStorage.Basket != "") {
                        var itemStrings = sessionStorage.Basket;
                        var sUserID = sessionStorage.UserID;
                        PageMethods.CreateBasketFromString(itemStrings, sUserID, OnSuccessCreateBasketFromString, OnErrorCreateBasketFromString);
                        ShowCheckOut();
                    }
                    else {

                        HideCheckOut();
                    }

                }
                else {

                    $.mobile.changePage("#pgLogin");

                }

                return false;

            });

            $("#btnBackProducts").click(function () {
                $.mobile.changePage("#pgLogin");
                ClearFormFields();
                ClearSession();
                return false;
            });

            $("#btnCheckOut").click(function () {


                if (sessionStorage.UserID != null) {
                    $.mobile.changePage("#pgAddress");
                    ClearFormFields();
                    PageMethods.GetCountries(OnSuccessGetCountries, OnErrorGetCountries);

                }
                else {

                    $.mobile.changePage("#pgLogin");
                    ClearFormFields();
                }

                return false;

            });

            $("#btnAddToBasket").click(function () {

                var qtyRequired = parseInt($("#txtQtyRequired").val());
                var qtyAvailable = parseInt($("#lblAvailableQty").html());

                if (qtyRequired > 0 && qtyRequired <= qtyAvailable) {

                    AddItemToBasket();
                    $.mobile.changePage("#pgProducts");
                    $("#txtQtyRequiredError").html("");
                    $("#txtQtyRequiredError").removeClass("Error");
                }
                else {

                    $("#txtQtyRequiredError").html("Quantity required should be greater than 0 and less than or equal to available quantity.");
                    $("#txtQtyRequiredError").addClass("Error");

                }


                return false;

            });

            $("#btnBackAddress").click(function () {

                $.mobile.changePage("#pgBasket");
                return false;

            });


            $("#btnFinish").click(function () {

                $.mobile.changePage("#pgLogin");
                return false;

            });

            $("#productsInBasket li .ui-li-link-alt").live("click", function () {

                var ID = $(this).attr('ID');
                $(this).parent("li").remove();
                RemoveItemFromBasket(ID);


            });

        });

        function HideCheckOut() {

            $("#divCheckOut").addClass("hide");
            $("#MsgBasket").html("Basket is empty.");
            $("#MsgBasket").addClass("Notification");

        }

        function ShowCheckOut() {

            $("#divCheckOut").removeClass("hide");
            $("#MsgBasket").html("");

        }

        function RemoveItemFromBasket(productKey) {

            if (sessionStorage.Basket != null) {

                if (sessionStorage.Basket.indexOf(productKey) != -1) {

                    var itemStrings = sessionStorage.Basket;
                    var items = itemStrings.split(/[\s,]+/);
                    if (items != null && items.length > 0) {
                        for (i = 0; i < items.length; i++) {

                            if (items[i] == productKey) {

                                items.splice(i, 2);
                                sessionStorage.Basket = items.toString();

                                if (sessionStorage.Basket == null || sessionStorage.Basket == "") {

                                    HideCheckOut();

                                }
                                else {

                                    ShowCheckOut();
                                }


                                break;
                            }
                        }
                    }

                }

            }

        }

        function ShowBasket(products) {

            $('#productsInBasket').empty();

            $.each(products, function (index, product) {

                var $list = $('#productsInBasket');
                $('<li />', { ID: product.LogisticProductKey })
                .append($('<a>'))
                .append($('<h2>', { text: product.Product, class: 'Basket' }))
                .append($('<p>', { text: product.ProductDescription, class: 'Basket' }))
                .append($('<span />', { text: product.Quantity, class: 'ui-li-count' }))
                .append($('<a>', { ID: product.LogisticProductKey }))
                .appendTo($list);
                $list.listview('refresh');

            });


        }


        function AddItemToBasket() {
            var qtyRequired = $("#txtQtyRequired").val();
            var productKey = sessionStorage.CurrentProductKeyInBasket;
            if (sessionStorage.Basket != null && sessionStorage.Basket != "") {

                if (sessionStorage.Basket.indexOf(productKey) != -1) {

                    UpdateQtyInBasket(productKey, qtyRequired);

                }
                else {

                    sessionStorage.Basket = sessionStorage.Basket + ',' + sessionStorage.CurrentProductKeyInBasket + ',' + qtyRequired;
                }


            }
            else {

                sessionStorage.Basket = sessionStorage.CurrentProductKeyInBasket + ',' + qtyRequired;
            }

        }

        function BindCountryList(data) {

            var countries = data;

            $('#selectCountry').empty();

            $.each(countries, function (index, country) {

                $('#selectCountry').append('<option value=' + country.CountryKey + '>' + country.CountryName + '</option>');

            });

            $('#selectCountry').refresh();

        }


        function OnSuccessSubmitConsignment(msg) {

            $.mobile.changePage("#pgFinish");
            $("#lblConsignmentNo").html("Your order has been placed succesfully. The consignment number of your order is " + msg);
            $("#lblConsignmentNo").addClass("Notification");
            ClearSession();
            ClearFormFields();
            ClearBasket();

        }

        function OnErrorSubmitConsignment(msg) {

            $("#lblConsignmentNo").html("some error has occured while processing your order. please try again by pressing the finish button.");
            $("#lblConsignmentNo").addClass("Error");
            ClearSession();
            ClearFormFields();
            ClearBasket();

        }

        function OnSuccessGetCountries(data) {

            BindCountryList(data);

        }

        function OnErrorGetCountries(msg) {

            alert(msg);

        }



        function OnSuccessCreateBasketFromString(data) {

            ShowBasket(data);

        }


        function OnErrorCreateBasketFromString(msg) {

            alert(msg);

        }




        function UpdateQtyInBasket(nLogisticProductKey, qty) {

            var itemStrings = sessionStorage.Basket;
            var items = itemStrings.split(/[\s,]+/);

            if (items != null && items.length > 0) {

                for (i = 0; i < items.length; i++) {

                    if (items[i] = nLogisticProductKey) {

                        var j = i + 1;
                        items[j] = qty;
                        sessionStorage.Basket = items.toString();
                        break;
                    }
                }
            }

        }

        function OnSuccessGetProdInfoByID(product) {

            $("#lblAvailableQty").html(product.Quantity);
            $("#lblProductCode").html(product.ProductDescription);
            $("#imgProduct").attr('src', 'http://my.transworld.eu.com/common/prod_images/thumbs/' + product.ThumbNailImage);

        }

        function OnErrorGetProdInfoByID(msg) {

            alert(msg);

        }

        function OnSuccessGetProducts(data) {

            $.mobile.changePage("#pgProducts");

            var products = data;

            var $list = $('#productList');

            $list.empty();

            $.each(products, function (index, product) {

                $('<li/>')
                .append($('<img>', { src: 'http://my.transworld.eu.com/common/prod_images/thumbs/' + product.ThumbNailImage, id: product.LogisticProductKey }))
                .append($('<h2>', { text: product.Product }))
                .append($('<p>', { text: product.ProductDescription }))
                .append($('<span />', { text: product.Quantity, class: 'ui-li-count' }))
                .appendTo($list);

            });

            $list.listview('refresh');

        }

        function OnErrorGetProducts(msg) {
            alert("OnErrorGetProducts" + msg.ID);
        }



        function OnSuccessVerifyPassword(msg) {            

            if (msg) {

                var isRemember = $("#selectRemeberMe").attr('value');
                var sUserID = $("#txtUserID").val();               
                
                if (isRemember == 'true') 
                {
                    localStorage.UserID = sUserID;
                }
                else 
                {
                    localStorage.UserID = null;
                }
                
                sessionStorage.UserID = sUserID.toUpperCase();
                PageMethods.GetProducts(sUserID, OnSuccessGetProducts, OnErrorGetProducts);
            }
            else {
                $("#ErrorUserIDError").addClass("Error");
                $("#ErrorUserIDError").html("User ID or Password doesn't match.");
            }

        }


        function OnErrorVerifyPassword(msg) {
            alert("error");
        }
    

    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:ScriptManager ID="sm" EnablePageMethods="true" runat="server">
        </asp:ScriptManager>
    </div>
    <div id="pgLogin" data-role="page">
        <div data-role="header" data-theme="b">
            <h1>
                Transworld Ordering System</h1>
        </div>
        <div data-role="content" data-theme="b">
            <div data-role="fieldcontain" data-theme="b">
                <div>
                    <em>*</em>User ID
                </div>
                <div>
                    <input id="txtUserID" type="text" maxlength="20" required />
                </div>
                <div class="clear">
                    &nbsp;</div>
                <div>
                    <em>*</em>Password
                </div>
                <div>
                    <input id="txtPassword" type="password" maxlength="20" required />
                </div>
                <div class="clear">
                </div>
                <div>
                    <label for="selectRemeberMe">
                        Remember me</label>
                </div>
                <div>
                    <select id="selectRemeberMe" data-role="slider">
                        <option value="false">Off</option>
                        <option value="true">On</option>
                    </select>
                </div>
                <div>
                    <span id="ErrorUserIDError"></span>
                </div>
                <div class="clear">
                    &nbsp;</div>
                <div>
                    <button id="btnLogin" type="submit" data-role="button" data-theme="b">
                        Login</button>
                </div>
            </div>
        </div>
    </div>
    <div id="pgProducts" data-role="page" data-dom-cache="true">
        <div data-role="header" data-theme="b" data-position="fixed">
            <a href="#" id="btnBackProducts">back</a>
            <h1>
                Products List</h1>
            <a href="#" id="btnShowBasket">show basket</a>
        </div>
        <div data-role="content" data-theme="b">
            <div data-role="fieldcontain">
                <div id="divProductList">
                    <ul id="productList" data-role="listview" data-inset="true" data-theme="b" data-filter='true'>
                    </ul>
                </div>
            </div>
        </div>
    </div>
    <div id="pgProductDetail" data-role="page">
        <div data-role="header" data-theme="b">
            <h1>
                Product Detail</h1>
        </div>
        <div data-role="content" data-theme="b">
            <div data-role="fieldcontain">
                <div class="ui-grid-a">
                    <div class="ui-block-a">
                        <img id="imgProduct" />
                    </div>
                    <div class="ui-block-b">
                        <label id="lblProductCode">
                        </label>
                    </div>
                </div>
                <div class="ui-grid-a">
                    <div class="ui-block-a">
                        Quantity Available
                    </div>
                    <div class="ui-block-b">
                        <label id="lblAvailableQty">
                        </label>
                    </div>
                </div>
                <div class="ui-grid-a">
                    <div class="ui-block-a">
                        Quantity Required
                    </div>
                    <div class="ui-block-b">
                        <input id="txtQtyRequired" type="number" min="1" maxlength="5" />
                        <span id="txtQtyRequiredError"></span>
                    </div>
                </div>
                <div class="ui-grid-a">
                    <div class="ui-block-a">
                        <button id="btnBackProductDetail" data-role="button">
                            Back</button>
                    </div>
                    <div class="ui-block-b">
                        <button id="btnAddToBasket" data-role="button" data-theme="b">
                            Add to basket</button>
                    </div>
                </div>
            </div>
            <div data-role="popup" id="divPopupImage" data-overlay-theme="a">
                <a href="#" data-rel="back" data-role="button" data-theme="a" data-icon="delete"
                    data-iconpos="notext" class="ui-btn-right">Close</a>
                <img id="imgProductPopup" alt="Photo Run" />
            </div>
        </div>
    </div>
    <div id="pgBasket" data-role="page">
        <div data-role="header" data-theme="b" data-position="fixed">
            <h1>
                Basket</h1>
            <a href="#" id="btnBackBasket">Back</a>
        </div>
        <div data-role="content" data-theme="b">
            <div data-role="fieldcontain">
                <div>
                    <ul id="productsInBasket" data-role="listview" data-inset="true" data-split-theme="d"
                        data-split-icon="delete">
                    </ul>
                </div>
                <div>
                    <span id="MsgBasket"></span>
                </div>
                <div id="divCheckOut">
                    <button id="btnCheckOut" data-role="button">
                        Check out</button>
                </div>
            </div>
        </div>
    </div>
    <div id="pgAddress" data-role="page">
        <div data-role="header" data-theme="b">
            <h1>
                Delivery Address</h1>
        </div>
        <div data-role="content" data-theme="b">
            <div data-role="fieldcontain">
                <div class="ui-grid-a">
                    <div class="ui-block-a">
                        Company Name
                    </div>
                    <div class="ui-block-b">
                        <input id="txtCompanyName" required="required" maxlength="50" />
                    </div>
                    <div class="ui-block-a">
                        &nbsp;
                    </div>
                    <div class="ui-block-b">
                        <label id="lblCompanyNameError">
                        </label>
                    </div>
                    <div class="ui-block-a">
                        Contact Name
                    </div>
                    <div class="ui-block-b">
                        <input id="txtContactName" required="required" maxlength="50" />
                    </div>
                    <div class="ui-block-a">
                        &nbsp;
                    </div>
                    <div class="ui-block-b">
                        <label id="lblContactNameError">
                        </label>
                    </div>
                    <div class="ui-block-a">
                        Address 1
                    </div>
                    <div class="ui-block-b">
                        <input id="txtAddress1" required="required" maxlength="50" />
                    </div>
                    <div class="ui-block-a">
                        &nbsp;
                    </div>
                    <div class="ui-block-b">
                        <label id="lblAddress1Error">
                        </label>
                    </div>
                    <div class="ui-block-a">
                        Address 2
                    </div>
                    <div class="ui-block-b">
                        <input id="txtAddress2" required="required" maxlength="50" />
                    </div>
                    <div class="ui-block-a">
                        &nbsp;
                    </div>
                    <div class="ui-block-b">
                        <label id="lblAddress2Error">
                        </label>
                    </div>
                    <div class="ui-block-a">
                        Town
                    </div>
                    <div class="ui-block-b">
                        <input id="txtTown" required="required" maxlength="50" />
                    </div>
                    <div class="ui-block-a">
                        &nbsp;
                    </div>
                    <div class="ui-block-b">
                        <label id="lblTownError">
                        </label>
                    </div>
                    <div class="ui-block-a">
                        Post Code
                    </div>
                    <div class="ui-block-b">
                        <input id="txtPostCode" required="required" maxlength="10" />
                    </div>
                    <div class="ui-block-a">
                        &nbsp;
                    </div>
                    <div class="ui-block-b">
                        <label id="lblPostCodeError">
                        </label>
                    </div>
                    <div class="ui-block-a">
                        Country
                    </div>
                    <div class="ui-block-b">
                        <select data-mini="true" id="selectCountry" data-theme="b">
                        </select>
                    </div>
                    <div class="ui-block-a">
                        &nbsp;
                    </div>
                    <div class="ui-block-b">
                        <label id="lblSelectCountryError">
                        </label>
                    </div>
                    <div class="ui-block-a">
                        Customer Reference
                    </div>
                    <div class="ui-block-b">
                        <input id="txtCustRef1" required="required" maxlength="50" />
                    </div>
                    <div class="ui-block-a">
                        &nbsp;
                    </div>
                    <div class="ui-block-b">
                        <label id="lblCustRef1Error">
                        </label>
                    </div>
                    <div class="ui-block-a">
                        Special Instructions
                    </div>
                    <div class="ui-block-b">
                        <input id="txtSpecialInstructions" required="required" maxlength="1000" />
                    </div>
                    <div class="ui-block-a">
                        &nbsp;
                    </div>
                    <div class="ui-block-b">
                        <label id="lblSpecialInstError">
                        </label>
                    </div>
                    <div class="ui-block-a">
                        <button id="btnBackAddress" data-role="button">
                            Back</button>
                    </div>
                    <div class="ui-block-b">
                        <button id="btnConfirmOrder" data-role="button">
                            Confirm Order</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <div id="pgFinish" data-role="page">
        <div data-role="header" data-theme="b">
            <h1>
                Order Confirmation</h1>
        </div>
        <div data-role="content" data-theme="b">
            <div>
                <label id="lblConsignmentNo">
                </label>
            </div>
            <div>
                <button id="btnFinish" data-role="button">
                    Finish</button>
            </div>
        </div>
    </div>
    </form>
</body>
</html>
<script runat="server">
    
    Private Shared gsConn As String = ConfigurationManager.ConnectionStrings("AIMSRootConnectionString").ToString
    Private Shared sProdThumbFolder As String = ConfigLib.GetConfigItem_prod_thumb_folder
    
    <System.Web.Services.WebMethod()>
    Public Shared Function nSubmitConsignment(ByVal sUserID As String, ByVal sCompanyName As String, ByVal sCtcName As String, ByVal sAddress1 As String, ByVal sAddress2 As String, ByVal sTown As String, ByVal sPostCode As String, ByVal sCustRef1 As String, ByVal sSpecialInst As String, ByVal nCountryKey As Integer, ByVal sProductsAndQty As String) As Integer
        
        If Not String.IsNullOrEmpty(sUserID) Then
            
            Dim sSQL As String = "select [Key] 'UserKey', CustomerKey from UserProfile where UserID = '" & sUserID.Replace("'", "''") & "'"
            Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
            
            If Not oDataTable Is Nothing AndAlso oDataTable.Rows.Count <> 0 Then
                
                Dim dr As DataRow = oDataTable.Rows(0)
                Dim nCustomerKey As Int32 = Convert.ToInt32(dr("CustomerKey"))
                Dim nUserKey As Int32 = Convert.ToInt32(dr("UserKey"))
                Dim lBookingKey As Long
                Dim lConsignmentKey As Long
                Dim BookingFailed As Boolean
                Dim oConn As New SqlConnection(gsConn)
                Dim oTrans As SqlTransaction
                Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
                
                nSubmitConsignment = 0
                oCmdAddBooking.CommandType = CommandType.StoredProcedure
    
                Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
                param1.Value = nUserKey
                oCmdAddBooking.Parameters.Add(param1)
        
                Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                param2.Value = nCustomerKey
                oCmdAddBooking.Parameters.Add(param2)
        
                Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
                param2a.Value = "MOBILE_BOOKING"
                oCmdAddBooking.Parameters.Add(param2a)
        
                Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
                Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
                Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
                Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)

                param3.Value = sCustRef1
                param4.Value = ""
                param5.Value = ""
                param6.Value = ""
        
       

                oCmdAddBooking.Parameters.Add(param3)
                oCmdAddBooking.Parameters.Add(param4)
                oCmdAddBooking.Parameters.Add(param5)
                oCmdAddBooking.Parameters.Add(param6)

                Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
                param6a.Value = Nothing
                oCmdAddBooking.Parameters.Add(param6a)
                
                Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
                param7.Value = sSpecialInst.Replace(Environment.NewLine, " ").Trim
                oCmdAddBooking.Parameters.Add(param7)
                Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
                param8.Value = ""
                oCmdAddBooking.Parameters.Add(param8)
                
                
                Dim param9 As SqlParameter = New SqlParameter("@ConsignmentType", SqlDbType.NVarChar, 20)
                param9.Value = "STOCK ITEM"
                oCmdAddBooking.Parameters.Add(param9)
                Dim param10 As SqlParameter = New SqlParameter("@ServiceLevelKey", SqlDbType.Int, 4)
                param10.Value = -1
                oCmdAddBooking.Parameters.Add(param10)
                Dim param11 As SqlParameter = New SqlParameter("@Description", SqlDbType.NVarChar, 250)
                param11.Value = "PRINTED MATTER - FREE DOMICILE"
                oCmdAddBooking.Parameters.Add(param11)

                Dim dtCnor As DataTable = ExecuteQueryToDataTable("SELECT * FROM Customer WHERE CustomerKey = " & nCustomerKey)
                Dim drCnor As DataRow
                If dtCnor.Rows.Count = 1 Then
                    drCnor = dtCnor.Rows(0)
                Else
                    WebMsgBox.Show("Couldn't find Consignor details.")
                    Exit Function
                End If
       
                Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
                'param13.Value = psCnorCompany
                param13.Value = drCnor("CustomerName")
       
                oCmdAddBooking.Parameters.Add(param13)
                Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
                param14.Value = drCnor("CustomerAddr1")
                oCmdAddBooking.Parameters.Add(param14)
                Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
                param15.Value = drCnor("CustomerAddr2")
                oCmdAddBooking.Parameters.Add(param15)
                Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
                param16.Value = drCnor("CustomerAddr3")
                oCmdAddBooking.Parameters.Add(param16)
                Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
                param17.Value = drCnor("CustomerTown")
                oCmdAddBooking.Parameters.Add(param17)
                Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
                param18.Value = drCnor("CustomerCounty")
                oCmdAddBooking.Parameters.Add(param18)
                Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
                param19.Value = drCnor("CustomerPostCode")
                oCmdAddBooking.Parameters.Add(param19)
                Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
                param20.Value = drCnor("CustomerCountryKey")
                oCmdAddBooking.Parameters.Add(param20)
                Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
                param21.Value = ""
                oCmdAddBooking.Parameters.Add(param21)
                Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
                param22.Value = ""
                oCmdAddBooking.Parameters.Add(param22)
                Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
                param23.Value = ""
                oCmdAddBooking.Parameters.Add(param23)
                Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
                param24.Value = 0
                oCmdAddBooking.Parameters.Add(param24)
                
                '''''' from textboxes
                
                Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
                param25.Value = sCompanyName
                oCmdAddBooking.Parameters.Add(param25)
                Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
                param26.Value = sAddress1
                oCmdAddBooking.Parameters.Add(param26)
                Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
                param27.Value = sAddress2
                oCmdAddBooking.Parameters.Add(param27)
                Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
                param28.Value = ""
                oCmdAddBooking.Parameters.Add(param28)
                Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
                param29.Value = sTown
                oCmdAddBooking.Parameters.Add(param29)
                Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
                param30.Value = ""
                oCmdAddBooking.Parameters.Add(param30)
                Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
                param31.Value = sPostCode
                oCmdAddBooking.Parameters.Add(param31)
                Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
                param32.Value = nCountryKey
                oCmdAddBooking.Parameters.Add(param32)
                Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
                param33.Value = sCtcName
                oCmdAddBooking.Parameters.Add(param33)
                Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
                param34.Value = ""
                oCmdAddBooking.Parameters.Add(param34)
                Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
                param35.Value = ""
                oCmdAddBooking.Parameters.Add(param35)
                Dim param36 As SqlParameter = New SqlParameter("@CneePreAlertFlag", SqlDbType.Bit)
                param36.Value = 0
                oCmdAddBooking.Parameters.Add(param36)
                Dim param37 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                param37.Direction = ParameterDirection.Output
                oCmdAddBooking.Parameters.Add(param37)
                Dim param38 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
                param38.Direction = ParameterDirection.Output
                oCmdAddBooking.Parameters.Add(param38)
        
                'For i As Int32 = 0 To oCmdAddBooking.Parameters.Count - 1
                '    Trace.Write(oCmdAddBooking.Parameters(i).ParameterName.ToString)
                '    Trace.Write(oCmdAddBooking.Parameters(i).DbType.ToString)
                '    If Not IsNothing(oCmdAddBooking.Parameters(i).Value) Then
                '        Trace.Write(oCmdAddBooking.Parameters(i).Value.ToString)
                '    Else
                '        Trace.Write("NOTHING")
                
                '    End If
                'Next
        
                Try
                    BookingFailed = False
                    oConn.Open()
                    oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddBooking")
                    oCmdAddBooking.Connection = oConn
                    oCmdAddBooking.Transaction = oTrans
                    oCmdAddBooking.ExecuteNonQuery()
                    lBookingKey = CLng(oCmdAddBooking.Parameters("@LogisticBookingKey").Value.ToString)
                    lConsignmentKey = CLng(oCmdAddBooking.Parameters("@ConsignmentKey").Value.ToString)
                    If lBookingKey > 0 Then
                        Dim IListOfProducts As List(Of Product) = CreateBasketFromString(sProductsAndQty, sUserID)
                        For Each product As Product In IListOfProducts
                            Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                            oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                            Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                            param51.Value = nUserKey
                            oCmdAddStockItem.Parameters.Add(param51)
                            Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                            param52.Value = nCustomerKey
                            oCmdAddStockItem.Parameters.Add(param52)
                            Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                            param53.Value = lBookingKey
                            oCmdAddStockItem.Parameters.Add(param53)
                            Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                            param54.Value = product.LogisticProductKey
                            oCmdAddStockItem.Parameters.Add(param54)
                            Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                            param55.Value = "PENDING"
                            oCmdAddStockItem.Parameters.Add(param55)
                            Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                            param56.Value = product.Quantity
                            oCmdAddStockItem.Parameters.Add(param56)
                            Dim param57 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 8)
                            param57.Value = lConsignmentKey
                            oCmdAddStockItem.Parameters.Add(param57)
                            oCmdAddStockItem.Connection = oConn
                            oCmdAddStockItem.Transaction = oTrans
                            oCmdAddStockItem.ExecuteNonQuery()
                        Next
                        Dim oCmdCompleteBooking As SqlCommand = New SqlCommand("spASPNET_LogisticBooking_Complete", oConn)
                        oCmdCompleteBooking.CommandType = CommandType.StoredProcedure
                        Dim param71 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                        param71.Value = lBookingKey
                        oCmdCompleteBooking.Parameters.Add(param71)
                        oCmdCompleteBooking.Connection = oConn
                        oCmdCompleteBooking.Transaction = oTrans
                        oCmdCompleteBooking.ExecuteNonQuery()
                    Else
                        BookingFailed = True
                    End If
                    If Not BookingFailed Then
                        oTrans.Commit()
                        nSubmitConsignment = lConsignmentKey
                    Else
                        oTrans.Rollback("AddBooking")
                    End If
                Catch ex As SqlException
                    oTrans.Rollback("AddBooking")
                Finally
                    oConn.Close()
                End Try
            
            End If
                
        End If
     
    End Function
    
    
    <System.Web.Services.WebMethod()>
    Public Shared Function GetCountries() As List(Of Country)
        Dim IListOfCountry As New List(Of Country)
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Country_GetCountries", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oConn.Open()
            Dim reader As SqlDataReader = oCmd.ExecuteReader()
            oDataTable.Load(reader)
            For Each dr As DataRow In oDataTable.Rows
                Dim country As New Country
                country.CountryName = dr("CountryName").ToString
                country.CountryKey = dr("CountryKey").ToString
                IListOfCountry.Add(country)
            Next
            GetCountries = IListOfCountry
        Catch ex As SqlException
            GetCountries = Nothing
        Finally
            oConn.Close()
        End Try
    End Function
    
    <System.Web.Services.WebMethod()>
    Public Shared Function VerifyUserCredentials(ByVal sUserID As String, ByVal sPassword As String) As Boolean
        
        VerifyUserCredentials = False
        
        Dim oUserInfo As SprintInternational.UserInfo = New SprintInternational.UserInfo()
        Dim oLogon As SprintInternational.Logon = New SprintInternational.Logon()
        Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
        
        oUserInfo = oLogon.GetUserInfo(sUserID)
        
        If oUserInfo.UserKey = -1 Then
            Exit Function
        Else
            Dim sActualPassword As String = oPassword.Decrypt(oUserInfo.Password)
            If sActualPassword = sPassword Then
                VerifyUserCredentials = True
            End If
        End If
        
    End Function
    
    <System.Web.Services.WebMethod()>
    Public Shared Function CreateBasketFromString(ByVal sProductsAndQty As String, ByVal sUserID As String) As List(Of Product)
        
        Dim items() As String = sProductsAndQty.Split(",")
        Dim IListOfProducts As New List(Of Product)
        
        For i = 0 To items.Length - 1 Step 2
            Dim sSQL As String = "select LogisticProductKey, ProductCode + ' ' + ISNULL(ProductDate,'') 'Product', ProductDescription, ThumbnailImage from LogisticProduct where LogisticProductKey = " & items(i)
            Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
            If Not oDataTable Is Nothing AndAlso oDataTable.Rows.Count > 0 Then
                For Each dr As DataRow In oDataTable.Rows
                    Dim objProduct As New Product
                    objProduct.LogisticProductKey = Convert.ToInt32(dr("LogisticProductKey"))
                    objProduct.Product = dr("Product").ToString()
                    objProduct.ProductDescription = dr("ProductDescription").ToString
                    objProduct.ThumbNailImage = dr("ThumbNailImage").ToString()
                    Dim j = i + 1
                    objProduct.Quantity = items(j)
                    IListOfProducts.Add(objProduct)
                Next
            End If
        Next
        
        CreateBasketFromString = IListOfProducts
        
    End Function
    
    <System.Web.Services.WebMethod()>
    Public Shared Function GetProducts(ByVal sUserID As String) As List(Of Product)
        
        GetProducts = Nothing
        Dim oDTProducts As New DataTable
        Dim IListOfProducts As New List(Of Product)
        Dim nUserKey As Int32
        Dim nCustomerKey As Int32
        Dim sSQL As String = "select [Key] 'UserKey', CustomerKey from UserProfile where UserId = '" & sUserID.ToUpper.Replace("'", "''") & "'"
        Dim oDTUserProfile As DataTable = ExecuteQueryToDataTable(sSQL)
        If Not oDTUserProfile Is Nothing AndAlso oDTUserProfile.Rows.Count > 0 Then
            For Each row As DataRow In oDTUserProfile.Rows
                nUserKey = Convert.ToInt32(row("UserKey"))
                nCustomerKey = Convert.ToInt32(row("CustomerKey"))
                Dim oConn As New SqlConnection(gsConn)
                Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_Mobile_GetProducts", oConn)
                oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                Dim IListOfParameters As New List(Of SqlParameter)
                Dim paramCustomer As New SqlParameter("@CustomerKey", SqlDbType.Int)
                paramCustomer.Value = nCustomerKey
                IListOfParameters.Add(paramCustomer)
                Dim paramUserKey As New SqlParameter("@UserKey", SqlDbType.Int)
                paramUserKey.Value = nUserKey
                IListOfParameters.Add(paramUserKey)
                oCmdAddStockItem.Parameters.AddRange(IListOfParameters.ToArray)
                Try
                    oConn.Open()
                    Dim reader As SqlDataReader = oCmdAddStockItem.ExecuteReader()
                    oDTProducts.Load(reader)
                    For Each dr As DataRow In oDTProducts.Rows
                        Dim objProduct As New Product
                        objProduct.LogisticProductKey = Convert.ToInt32(dr("LogisticProductKey"))
                        objProduct.Product = dr("Product").ToString()
                        objProduct.ProductDescription = dr("ProductDescription").ToString()
                        objProduct.ThumbNailImage = dr("ThumbNailImage").ToString()
                        objProduct.Quantity = Convert.ToInt32(dr("Quantity"))
                        IListOfProducts.Add(objProduct)
                    Next
                Catch ex As SqlException
                    GetProducts = Nothing
                Finally
                    oConn.Close()
                End Try
            Next
        End If
        GetProducts = IListOfProducts
        
    End Function
    
    <System.Web.Services.WebMethod()>
    Public Shared Function GetProdInfoByID(ByVal sLogisticProductKey As String) As Product
        Dim oDataTable As New DataTable
        Dim sSQL As String = "SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ")) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ") END, "
        sSQL += "(SELECT ProductCode + ' ' + ISNULL(ProductDate,'') + ' ' + ProductDescription from LogisticProduct WHERE LogisticProductKey = " & sLogisticProductKey & ")'ProdDescription', "
        sSQL += "(SELECT ThumbNailImage from LogisticProduct where LogisticProductKey = " & sLogisticProductKey & ")'ThumbNailImage'"
        Try
            oDataTable = ExecuteQueryToDataTable(sSQL)
            Dim row As DataRow = oDataTable.Rows(0)
            Dim oProduct As New Product
            oProduct.ProductDescription = row("ProdDescription").ToString
            oProduct.Quantity = Convert.ToInt32(row("Quantity"))
            oProduct.ThumbNailImage = row("ThumbNailImage").ToString
            GetProdInfoByID = oProduct
        Catch
            GetProdInfoByID = Nothing
        End Try
    End Function
    
    '<System.Web.Services.WebMethod()>
    'Public Shared Function GetProductCode(ByVal sLogisticProductKey As String) As String
    '    Dim sSQL As String = "select ProductCode + ' ' + ISNULL(ProductDate,'') + ' ' + ProductDescription 'ProdDescription' from LogisticProduct WHERE LogisticProductKey = " & sLogisticProductKey
    '    Try
    '        GetProductCode = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    '    Catch
    '        GetProductCode = String.Empty
    '    End Try
    'End Function
    
    
    <System.Web.Services.WebMethod()>
    Public Shared Function GetProductKeyByProductCode(ByVal sLogisticProductKey As String) As String
        Dim sSQL As String = "select LogisticProductKey from LogisticProduct WHERE ProductCode = '" & sLogisticProductKey & "'"
        Try
            GetProductKeyByProductCode = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
        Catch
            GetProductKeyByProductCode = String.Empty
        End Try
    End Function
    
    Private Shared Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sQuery, oConn)
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Try
            oAdapter.Fill(oDataTable)
            oConn.Open()
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTable = oDataTable
    End Function
    
    Public Class Country
        
        Private m_CountryName As String
        Private m_CountryKey As Int32
        
        Public Property CountryName() As String
            Get
                Return m_CountryName
            End Get
            Set(value As String)
                m_CountryName = value
            End Set
        End Property
        
        Public Property CountryKey() As Int32
            Get
                Return m_CountryKey
            End Get
            Set(value As Int32)
                m_CountryKey = value
            End Set
        End Property
        
    End Class
    
    Public Class Product
        
        Private m_LogisticProductKey As Int32
        Private m_quantity As Int32
        Private m_Product As String
        Private m_ProductDescription As String
        Private m_ThumbNailImage As String
        
        Public Property LogisticProductKey() As Int32
            Get
                Return m_LogisticProductKey
            End Get
            Set(value As Int32)
                m_LogisticProductKey = value
            End Set
        End Property
        
        Public Property Product() As String
            Get
                Return m_Product
            End Get
            Set(value As String)
                m_Product = value
            End Set
        End Property
        
        Public Property ProductDescription() As String
            Get
                Return m_ProductDescription
            End Get
            Set(value As String)
                m_ProductDescription = value
            End Set
        End Property
        
        Public Property ThumbNailImage() As String
            Get
                Return m_ThumbNailImage
            End Get
            Set(value As String)
                m_ThumbNailImage = value
            End Set
        End Property
        
        Public Property Quantity() As Int32
            Get
                Return m_quantity
            End Get
            Set(value As Int32)
                m_quantity = value
            End Set
        End Property
        
    End Class
    
    
</script>
