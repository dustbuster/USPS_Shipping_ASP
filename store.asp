
<%
'---------------------------------- USPS SHIPPING ---------------------------------------------'
' Most of these functions call each other, but it all originates in /includes/store_submit.asp '
'                                 Dustin H: Case 27993                                         '
'----------------------------------------------------------------------------------------------'
'CASES: TLS - 2073, 71, 92, 398, 417, 498, 529, 600 : Yeah... It got a bit cray cray getting this right

const USPSID = "YOUR_USPS_ID"

Function ShippingCost()
    'Determines if rate is specified by client or uses USPS API, Then sets Shipping costs
    dim uspsrate
    'Checks for international
    if request("s_country") <> "USA" then
        if AllItemsFreeShippingAndInt() then
            uspsRate = 0
        else
            uspsRate = 25
        End if
        ShippingCost = uspsRate
        session("AllItemsFreeShippingAndInt Function") = AllItemsFreeShippingAndInt()
        session("USPSRATE") = uspsrate
        Exit Function
    end if
    
    Select Case lcase(application("consoleFolderName"))
        case "inemi"
            if session("sa_id") <> "" then
                uspsRate = 0.00
            else 
                uspsRate = SendXMLGetShippingPrice()
            end if
        case else
            uspsRate = SendXMLGetShippingPrice()
    End Select
    ShippingCost = uspsRate
End Function

Function AllItemsFreeShippingAndInt()
    'This function pulls the total number of items
    dim freeCount, cartID, cartCount 
    cartID = session("cartid")
    Set freeCount = getrecordset("Select COUNT(*) "&_
        "from store_cart_items "&_
        "as sci "&_
        "join store_items "&_
        "as si on sci.itemid = si.id "&_
        "where si.freeship <> 0 "&_
        "and sci.cartid =" & cartID)
    set cartCount = getrecordset("select count(*) "&_
        "from store_cart_items "&_
        "where cartid ="& cartID)

    if freeCount.eof then
        AllItemsFreeShippingAndInt = false
        Exit Function
    End if
    session("TOTAL # of items:") = cartCount(0)
    session("TOTAL # of FREE items:") = freeCount(0)
    if cartCount(0) > freeCount(0) then
        AllItemsFreeShippingAndInt = false
    Else
        AllItemsFreeShippingAndInt = true
    End if
End Function



Function SendXMLGetShippingPrice()
    dim xml
    xml = BuildXml(packageCnt)
    if xml = "Error" then exit function
    ' Checks to see if all items have 0 lb and oz and exits with 0 cost shipping
    if packageCnt = 0 then
        SendXMLGetShippingPrice = 0
        exit function
    End if
    'Create Obj to parse xml response
    Set XHR = CreateObject("MSXML2.ServerXMLHTTP.3.0")
    XHR.Open "POST", "http://production.shippingapis.com/ShippingApi.dll?API=RateV4&", False
    XHR.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XHR.Send("API=RateV4&XML=" & Server.URLEncode(xml))
    'Sent!
    If XHR.Status <> 200 Then
        Response.write XHR.responseText
        SendXMLGetShippingPrice = "XML was not sent, please try again or contact site admin"
        Exit function
    End If

    session("XMLResponse2") = XHR.responseXML.XML
    dim cost, node, allNodes, errortxt, errorNodes
    Set allNodes = XHR.responseXML.selectNodes("/RateV4Response/Package/Postage/Rate")
    Set errorNodes = XHR.responseXML.selectNodes("/RateV4Response/Package/Error/Description")

    if errorNodes.length = 0 and allNodes.length = 0 then
        set errorNodes = XHR.responseXML.selectNodes("/Error/Description")
    end if

    If allNodes.length <> 0 and errorNodes.length = 0 then
        for each node in allNodes
            cost = cost + ccur(node.text)
        Next
        SendXMLGetShippingPrice = cost
    else
        for each node in errorNodes
            if node.text = "The element 'RateV4Request' has invalid child element 'Package'." then
                errortxt = errortxt & " - " & " You cannot ship more than 25 items in 1 order."
            else
                errortxt = errortxt & " - " & + node.text
            end if
        Next
        SendXMLGetShippingPrice = errortxt
    End if
End Function


Function BuildXml(packageCnt)
    set cartrs=getrecordset("SELECT store_cart_items.*, store_items.name, isNull(store_items.length, 0) as length, "&_
        "isNull(store_items.width, 0) as width, isnull(store_items.height, 0) as height, "&_ 
        "isnull(store_items.weight_oz, 0) as weight_oz, isNull(store_items.weight_lb,0) as weight_lb,"&_ 
        "store_items.tax_rate, store_items.media, isNull(store_items.itemnumber, 0) as itemnumber, store_items.freeship "&_
        "from store_cart_items "&_
        "join store_items on store_cart_items.itemid=store_items.id "&_ 
        "where store_items.freeship<>1 and cartid="&session("cartid"))

    dim container, size, mailService, girth, xml, shipFromZip, shipToZip, leng, wid, hei,lb,oz, freeShip
    container = "Variable" : size = "" : mailService = ""
    'Set Vars outside of loop
    shipFromZip = GetShipFromZipCode()
    shipToZip = request("s_zip")
    shipToZip = NineZipFormatter(shipToZip)
    'Start Building XML
    xml = "<?xml version=1.0 encoding=UTF-8?>"
    xml = xml & "<RateV4Request USERID=" & "'" & USPSID & "'" & ">"
    'Loops thru each package and determines the shipping price individually
    dim i : i = 0 : packageCnt = 0
    if Ucase(request("s_country")) = "USA" or lcase(request("s_country") = "usa") then
        do until cartrs.eof
            quantity = cartRs("qty")
            for i = 1 To quantity
                'Sanitize Values
                if cartrs("media") = false or isNull(cartrs("media")) then
                    mediaShip = false
                else
                    mediaShip = true
                End if

                'Check to see if free shippng is enable
                if cartrs("freeship") = false or isNull(cartrs("freeship")) then
                    freeShip = false
                else
                    freeShip = true
                End if

                'Dustinh, set vars for dynamic shipping
                lb = cint(cartrs("weight_lb"))
                oz = cint(cartrs("weight_oz"))
                if oz > 16 then
                    lb = lb + (oz / 16)
                    oz = oz mod 16
                end if
                leng = cint(cartrs("length"))
                wid = cint(cartrs("width"))
                hei = cint(cartrs("height"))
                'Check to see if weight is included
                if lb + oz > 0 or freeship = true then
                    packageCnt = packageCnt + 1
                    mailService = determineService(lb,oz,mediaShip)
                    size = determineSize(leng,wid,hei)
                    container = ContainerSize(size)
                    xml = xml & "<Package ID=" & "'" & packageCnt & "'" &">"
                    xml = xml & "<Service>" & mailService & "</Service>"
                    if mailService = "First Class" then
                        xml = xml & "<FirstClassMailType>Parcel</FirstClassMailType>"
                    end if
                    xml = xml & "<ZipOrigination>" & shipFromZip & "</ZipOrigination>"
                    xml = xml & "<ZipDestination>" & shipToZip & "</ZipDestination>"
                    xml = xml & "<Pounds>" & lb &  "</Pounds>"
                    xml = xml & "<Ounces>" & oz &  "</Ounces>"
                    xml = xml & "<Container>" & container & "</Container>"
                    xml = xml & "<Size>" & size & "</Size>"
                    if size = "Large" then
                        xml = xml & "<Width>" & wid & "</Width>"
                        xml = xml & "<Length>" & leng & "</Length>"
                        xml = xml & "<Height>" & hei & "</Height>"
                        xml = xml & "<Girth>" & girthCalculator(leng,wid,hei) & "</Girth>"
                    end if
                    xml = xml & "</Package>"
                end if
            Next
            cartrs.movenext
        Loop
    End if
    xml = xml & "</RateV4Request>"
    session("Built_XML") = xml
    BuildXml = xml
End Function

Function ContainerSize(size)
    if size = "Large" then
        ContainerSize = "Rectangular"
    else
        ContainerSize = "Variable"
    End if
End Function

Function GetShipFromZipCode()
    dim shipzip
    'This checks to see if the ShipFromZip is set in user defined text, 
    'if not it takes the most common zip from the admin users
    dim rsZip
    set rsZip = globalConnection.execute("select coalesce("&_
        "(select content from contentmodules where title='Ship from zip'),"&_
        "(select top 1 zip from cusers group by zip order by count(*) desc))")
    If rsZip.eof then
        GetShipFromZipCode="Error"
        Exit Function
    End if
    shipzip=StripTags(rsZip(0))
    shipzip=StripCharsSpaces(shipzip)
    GetShipFromZipCode = shipzip
End Function

Function NineZipFormatter(zip)
    'Simple ver
    result = ""
    if len(zip) => 9 then
        result = left(zip,5)
    else
        result = zip
    end if
    NineZipFormatter = result
End Function

Function StripTags(inputst)
    dim tempstr
    Dim regEx : Set regEx = New RegExp
    regEx.IgnoreCase = True : regEx.Global = True
    regEx.Pattern = "<.*?>"
    tempstr = regEx.Replace(inputst, "")
    StripTags=tempstr
End Function

Function StripCharsSpaces(inputSt)
    'Sanity checks the zip value itself after tags are removed
    dim tempstr : dim regEx : Set regEx = New RegExp
    regEx.IgnoreCase = True : regEx.Global = True
    'Checks to see if string is longer than index 5,
    if len(inputSt) > 5 then
        regEx.Pattern = "([\s]+)"  'removes spaces
        inputSt = regEx.Replace(inputSt, "")
        'checks for a dash, or if it's exactly 9 chars long
        if Instr(inputSt,"-") = 6 or len(inputSt) = 9 then 'Checks for 9 digit zip
            inputSt=left(inputSt,5)
        end if
    end if
    StripCharsSpaces = inputSt
End Function

Function determineSize(l,w,h)
    Dim size
    if l>12 or w>12 or h>12 then
        size="Large"
    else
        size="Regular"
    end if
    determineSize = size
End Function

Function determineService(lb,oz,mediaShip)
    service=""
    if lb =< 0 and oz < 6 and mediaShip <> true then
        service = "First Class"
    elseif mediaShip = true then
        service = "Media"
    elseif oz => 6 and lb => 0 and lb < 69 then
        service = "Priority"
    elseif lb < 69 then
        service = "Priority"
    else
        service = "Priority"
    end if
    determineService = service
End Function

Function girthCalculator(l,w,h)
    dim largest,g
    if l > w and l > h then
        largest = l
    elseif w > l and w > h then
        largest = w
    else
        largest = h
    End If
    g = 2 * ((l + w + h) - largest)
    girthCalculator = g
End Function


'- - - - - - - - - - - - - - - - - Functions and subs for store_submit.asp not shipping - - - - - - - - - - - - - - - - - -'

sub paybycheck(coid)
    globalconnection.execute("Update store_completedorders set paymentstatus='pending_inv' where id="&clng(coid))
    globalconnection.execute("Update store_completedorders set ordercomplete=1 where id="&clng(coid))
end sub


sub generateinv(coid,creditamount)
    'CALCULATE INVOICE CONTACTS/ORGS
    if session("sa_active")=1 and session("sa_suspension")=false and session("sa_isexpired")=false then
        'active member and members-only checkout
        contactid=session("sa_id")
        orgid=session("sa_memberid")
    else 'non logged in user
        contactid=clng(application("billme_nonmembers_inv_contactid"))
        orgid=0
        'lookup orgid
        dim olrs
        set olrs=getrecordset("select * from af_membermap where userid="&contactid)
        if not olrs.eof then
            orgid=olrs("memberid")
        end if
    end if
    'UPDATE STORE PURCHASE TO INDICATE PENDING INVOICE
    globalconnection.execute("Update store_completedorders set paymentstatus='pending_inv' where id="&clng(coid))
    globalconnection.execute("Update store_completedorders set ordercomplete=1 where id="&clng(coid))

    'GET STORE PURCHASE SUMMARY INFORMATION
    dim sprs
    set sprs=getrecordset("Select * from store_completedorders where id="&clng(coid))

    'GENERATE INVOICE SHELL
    dim irs
    set irs=server.CreateObject("ADODB.Recordset")
    irs.open "Select top 0 * from invoices", globalconnection, adopenstatic, adlockoptimistic
        irs.addnew
        irs("organizationid")=orgid
        irs("contactid")=contactid
        irs.update
    irs.close
    dim StoreGLRS
    set StoreGLRS = getrecordset("select GLCode from store_glcode")

    if not StoreGLRS.EOF then
        StoreGL = storeGLRS("GLCode")
    else
        storeGL = ""
    END IF

    dim mirs
    set mirs=getrecordset("Select max(id) as maxid from invoices where organizationid="&orgid&" and contactid="&contactid)
    invoiceid=mirs("maxid")
    'REOPEN INVOICE SHELL - POPULATE BASIC INFORMATION
    irs.open "Select * from invoices where id="&mirs("maxid"), globalconnection, adopenstatic, adlockoptimistic
        irs("fname")=trim(sprs("firstname"))
        irs("lname")=trim(sprs("lastname"))
        irs("address1")=trim(sprs("address1"))
        irs("address2")=trim(sprs("address2"))
        irs("city")=trim(sprs("city"))
        irs("state")=trim(sprs("state"))
        irs("zip")=trim(sprs("zip"))
        irs("country")=trim(sprs("country"))
        irs("email")=trim(sprs("email"))
        irs("termsid")=1
        irs("itemtotal")=sprs("itemtotal")+sprs("shippingtotal")
        irs("taxtotal")=trim(sprs("taxtotal"))
        irs("grandtotal")=sprs("itemtotal")+sprs("shippingtotal")+sprs("taxtotal")-promo_discount-creditamount
        irs("amountpaid")=0
        irs("createddate")=formatdatetime(now(), 2)
        irs("createdby")="Auto Generated"
        irs("storeid")=coid
        irs("notes")="Order #"&coid
        irs("posteddate")=Now()
        irs("postedstatus")=1
        irs("GLCode") = StoreGL
    irs.update
    irs.close

    'GET ALL CART ITEMS
    dim cirs
    set cirs=getrecordset("SELECT store_cart_items.*, store_items.name from store_cart_items join store_items on (store_cart_items.itemid=store_items.id) where cartid="&sprs("cartid"))

    'GET INVOICE LINE ITEM TYPE
    dim litype  '23173
    Select Case lcase(application("consolefoldername"))
        case "upceanet","inemi"
            litype=12
        case else
            litype=11
    End Select

    'ADD ALL INVOICE LINE ITEMS
    dim lirs
    set lirs=server.CreateObject("ADODB.Recordset")
    lirs.open "Select top 0 * from invoice_items", globalconnection, adopenstatic, adlockoptimistic
        do until cirs.eof
            lirs.addnew
            lirs("invoiceid")=invoiceid
            lirs("itemtypeid")=litype
            lirs("Description")=left(trim(cirs("name"))&" [Item# "&itemid&"] "&trim(cirs("options")),255)
            lirs("qty")=cirs("qty")
            lirs("unitprice")=cirs("cost")
            lirs("itemtotal")=cirs("qty")*cirs("cost")
            lirs.update
            cirs.movenext
        loop
        'ADD SHIPPING LINE ITEM
        if sprs("shippingtotal")>0 then
            lirs.addnew
            lirs("invoiceid")=invoiceid
            lirs("itemtypeid")=litype
            lirs("Description")="Shipping & Handling"
            lirs("qty")=1
            lirs("unitprice")=sprs("shippingtotal")
            lirs("itemtotal")=sprs("shippingtotal")
            lirs.update
       end if

        'ADD PROMO CODE DISCOUNT
        if promo_discount>0 then
            lirs.addnew
            lirs("invoiceid")=invoiceid
            lirs("itemtypeid")=litype
            lirs("Description")="Promotional Code Discount"
            lirs("qty")=1
            lirs("unitprice")=promo_discount*-1
            lirs("itemtotal")=promo_discount*-1
            lirs.update
        end if
        'ADD CREDIT DISCOUNT
        if creditamount>0 then
            lirs.addnew
            lirs("invoiceid")=invoiceid
            lirs("itemtypeid")=litype
            lirs("Description")="Account Credit Discount"
            lirs("qty")=1
            lirs("unitprice")=creditamount*-1
            lirs("itemtotal")=creditamount*-1
            lirs.update
        end if
    lirs.close
end sub


Function generatekey(storeitem, expirehours, expiredownloads, orderid)

    dim keyfound
    keyfound=false
    do until keyfound
        newkey=generatePassword(30)
        'check to see if this is in use already
        set rs=getrecordset("Select * from Store_downloadkeys where downloadkey='"&newkey&"'")
        if rs.eof then
            keyfound=true
        end if
    loop
    'newkey holds the contets of the new key - enter into database with parameters
    dim nrs
    set nrs=server.CreateObject("ADODB.Recordset")
    nrs.open "Select top 0 * from store_downloadkeys", globalconnection, adopenstatic, adlockoptimistic
    nrs.addnew
        nrs("itemid")=clng(storeitem)
        nrs("downloadkey")=newkey
        nrs("generated")=now()
        nrs("expirationdate")=dateadd("h", expirehours, now())
        nrs("totalattempts")=expiredownloads
        nrs("remainingattempts")=expiredownloads
        nrs("orderid")=clng(orderid)
    nrs.update
    nrs.close
    generatekey=newkey
end function


function df(inputtext)
    if len(inputtext)>0 then
        newinputtext=trim(inputtext)
        newinputtext=replace(newinputtext, chr(34), "&quot;")
        newinputtext=replace(newinputtext, "<", "&lt;")
        newinputtext=replace(newinputtext, ">", "&gt;")
        df=newinputtext
    else
        dispfield=""
    end if
end function


Function generatePassword(passwordLength)

    Dim sDefaultChars
    Dim iCounter
    Dim sMyPassword
    Dim iPickedChar
    Dim iDefaultCharactersLength
    Dim iPasswordLength
    sDefaultChars="abcdefghijklmnopqrstuvxyzABCDEFGHIJKLMNOPQRSTUVXYZ0123456789"
    iPasswordLength=passwordLength
    iDefaultCharactersLength = Len(sDefaultChars)
    Randomize
    For iCounter = 1 To iPasswordLength
        iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1)
        sMyPassword = sMyPassword & Mid(sDefaultChars,iPickedChar,1)
    Next
    generatePassword = sMyPassword
End Function

%>
