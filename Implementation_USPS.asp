<%

'So at this point where we enter the USPS code, a session variable of the "cartid" is active. 
'We call the function, the and functions grab the cart ID and we don't need to pass anything in. 

temp = ShippingCost()
if isNumeric(temp) then
    rs("shippingTotal") = temp
    session("itemShip") = temp
else
    'Error message after redirect
    Response.Redirect "../store_checkout.asp?err=true&msg=" & temp
end if
%>