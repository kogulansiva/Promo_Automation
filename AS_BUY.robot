*** Settings ***
Library           Selenium2Library
Resource          ReadExcel.py
Library           ReadExcel.ExcelUtility
Resource          site_Elements.txt

*** Variables ***
${firstName}      id=firstName    #First Name
${lastName}       id=lastName    #Last Name
${memberId}       id=memberId    #Member ID
${email}          id=email    #Email
${bgt-login-submit}    xpath=html/body/div[2]/div[1]/div/form/div[2]    #Submit Button
${bgt-offer-dropdown}    name=quantity    #Miles Offer Drop Down List
${introCopy}      xpath=html/body/div/div/div[2]/section/div/div/div/div[1]/div    #Intro Copy block
${legalCopy}      xpath=html/body/div/div/div[2]/section/div/div/div/div[2]/div[2]/p    #Legal Copy
${LCP_email}      xpath=html/body/div[2]/div/section/form/div[1]/input    #LCP login email address
${LCP_Password}    xpath=html/body/div[2]/div/section/form/div[2]/input    #LCP login Password
${LCP_Login_button}    xpath=html/body/div[2]/div/section/form/div[3]/button    #LCP Login Button
${Google_login_email}    id=Email    #Gmail email address
${Google_login_next_button}    id=next    # Google login next button
${Google_PW}      id=Passwd    # Google PW
${Google_signIn_Button}    id=signIn    #Google Sign-In Button
${LCP_admin_page}    id=app-container    #LCP admin page
${LCP_offer_tab}    xpath=html/body/div[1]/nav/ul/li/a[text()="Offers"]    # LCP Offer tab
${LCP_OfferName_Field}    xpath=html/body/div[2]/div/section/div[2]/div[2]/form/div[1]/div[1]/input    #LCP Offer name Search Input Field
${LCP_Search_button}    xpath=html/body/div[2]/div/section/div[2]/div[2]/form/div[1]/div[4]/div/div[2]/button    #LCP Offer Search Button
${LCP_Open_Offer}    xpath=html/body/div[2]/div/section/div[2]/div[2]/div[2]/table/tbody/tr/td[1]/a    #Open the offer from the seach results
${LCP_Offer_Base_PIC}    xpath=html/body/div[2]/div/section/div[2]/div/section[1]/table[2]/tbody/tr/td    #LCP Offer Base PIC
${LCP_Offer_Start_Date}    xpath=html/body/div[2]/div/section/div[2]/div/section[2]/table/tbody/tr[3]/td/table/tbody/tr/td[1]    #LCP offer start date
${LCP_Offer_End_Date}    xpath=html/body/div[2]/div/section/div[2]/div/section[2]/table/tbody/tr[3]/td/table/tbody/tr/td[2]    #LCP Offer End Date
${LCP_Offer_Type}    xpath=html/body/div[2]/div/section/div[2]/div/section[1]/table[1]/tbody/tr[3]/td    #LCP Offer Type
${LCP_Promo_Type}    xpath=html/body/div[2]/div/section/div[2]/div/section[3]/table[7]/tbody/tr[1]/td    #LCP Offer Promo Type
${LCP_Offer_Rate_Block_Size}    xpath=html/body/div[2]/div/section/div[2]/div/section[3]/table[8]/tbody/tr/td[1]    #LCP Offer Rate Blcok Size
${LCP_Offer_Rate_Effective}    xpath=html/body/div[2]/div/section/div[2]/div/section[3]/table[8]/tbody/tr/td[7]/table/tbody/tr[2]/td[2]    #LCP Offer effective rate
${LCP_Offer_Rate_Wholesale}    xpath=html/body/div[2]/div/section/div[2]/div/section[3]/table[8]/tbody/tr/td[8]/table/tbody/tr[2]/td[1]    #LCP offer wholesale rate
${LCP_Offer_Rate_Bonus_Rate}    xpath=html/body/div[2]/div/section/div[2]/div/section[3]/table[8]/tbody/tr/td[8]/table/tbody/tr[2]/td[2]    #LCP Offer Bonus Rate
${LCP_Offer_Rate_Commision_Rate}    xpath=html/body/div[2]/div/section/div[2]/div/section[3]/table[8]/tbody/tr/td[8]/table/tbody/tr[2]/td[3]    #LCP Offer Commission rate
${LCP_Offer_LPID}    xpath=html/body/div[2]/div/section/div[2]/div/section[5]/table/tbody/tr[1]/td[4]/strong    #LCP Offer LPID
${LCP_Offer_Member_List_Count}    xpath=html/body/div[2]/div/section/div[2]/div/section[5]/div/table/tbody/tr[2]/td    #LCP Offer Member Count
${LCP_Offer_Tags}    xpath=html/body/div[2]/div/section/div[2]/div/section[7]/div    #LCP Offer Tags
${LCP_Offer_Preview_Button}    xpath=html/body/div[2]/div/section/div[2]/div/section[4]/table/tbody/tr/td[2]/div    #LCP Offer Preview Button
${LCP_Offer_Priority}    xpath=html/body/div[2]/div/section/div[2]/div/section[2]/table/tbody/tr[4]/td    #LCP Offer Priority
${LCP_Preview_Storefront_Name}    name=storefrontLpName    #LCP Offer Preview form Storefront name
${LCP_Preview_Storefront_Type}    name=storefrontType    #LCP Offer Storefront type
${LCP_Preview_form_previewButton}    xpath=html/body/div[4]/div/form/div/button[2]    #LCP Offer Preview form preview button
${LCP_Offer_Tab_Production}    xpath=html/body/div[1]/nav/ul/li/a[text()="Offers"]    #LCP Offers tab in Production
${SecondTab}      url=https://storefront-staging.lxc.points.com/mileage-plan/en-US/buy    #Storefront Preview Page
${SF_UserName}    xpath=html/body/div[1]/div[1]/div/div/div[3]/div[3]/form/div[2]/div/input[1]
${SF_PW}          xpath=html/body/div[1]/div[1]/div/div/div[3]/div[3]/form/input[2]
${SF_Login_Button}    xpath=html/body/div[1]/div[1]/div/div/div[3]/div[3]/form/input[3]
${SF_Verify_Code}    xpath=html/body/div[1]/div/div/div[3]/form/div[1]/input
${SF_Verify_Code_Button}    xpath=html/body/div[1]/div/div/div[3]/form/input[10]
${StoreFront_Email}    id=billingEmail
${StoreFront_Phone}    id=phone
${StoreFront_Terms_CheckBox}    id=termsAndConditions
${StoreFront_Points_Logo_Legal}    xpath=html/body/div/div/div[2]/div/span
${Login_Box_Header}    xpath=html/body/div[2]/div[1]/div/h2
${LCP_Offer_Edit_Button}    xpath=html/body/div[2]/div/section/div[2]/div/section[8]/span[1]/a[text()="Edit"]
${LCP_Console_Offer_IntroCopy_Area}    xpath=html/body/div[2]/div/section/div[2]/div/form/div[1]/div[4]/div/div/div/div/div[3]/textarea
${LCP_Offer_Cancel_Button}    xpath=html/body/div[2]/div/section/div[2]/div/form/div[2]/a

*** Test Cases ***
TC3- Process A- PROD
    [Documentation]    *The test case will validate following contents of the offer in LCP Console:*
    ...
    ...    Start Date
    ...    End Date
    ...    Priority
    ...    PICs
    ...    Tags
    ...    Type of an offer(Buy or Gift, Bonus or Discount)
    ...    Rates
    ...    Lists Count
    Open Browser    https://admin.lcp.points.com/login    Chrome    #Production
    Maximize browser window
    Wait Until Element Is Visible    ${LCP_email}    20s
    Input Text    ${LCP_email}    kogulan.siva@points.com
    Input Text    ${LCP_Password}    Goby@india
    Click Element    ${LCP_Login_button}
    Comment    Input Text    ${Google_login_email}    kogulan.siva@points.com
    Comment    Click Element    ${Google_login_next_button}
    Comment    sleep    1s
    Comment    Input Text    ${Google_PW}    Goby@india
    Comment    Click Element    ${Google_signIn_Button}
    Wait Until Element Is Visible    ${LCP_Offer_Tab_Production}    8s
    Click Element    ${LCP_Offer_Tab_Production}
    Sleep    5s
    Input Text    ${LCP_OfferName_Field}    Alaska Buy - Personalized Storefront 25K BASE OFFER - Canada - TEST V1
    Click element    ${LCP_Search_button}
    Wait Until Element Is Visible    ${LCP_Open_Offer}    40s
    Click element    ${LCP_Open_Offer}
    Wait Until Element Is Visible    ${LCP_Offer_Type}
    Element Text Should Be    ${LCP_Offer_Type}    BUY
    Element Text Should Be    ${LCP_Offer_Base_PIC}    Points.com Instant Points
    Element Text Should Be    ${LCP_Offer_Start_Date}    26/10/2016 09:00:00 EDT
    Element Text Should Be    ${LCP_Offer_End_Date}    01/04/2017 02:59:00 EDT
    Element Text Should Be    ${LCP_Offer_Priority}    8000
    Element Text Should Be    ${LCP_Promo_Type}    none
    Element Text Should Be    ${LCP_Offer_Rate_Block_Size}    1,000
    Element Text Should Be    ${LCP_Offer_Rate_Effective}    0.0275
    Element Text Should Be    ${LCP_Offer_Rate_Wholesale}    0.0252 per point
    Element Text Should Be    ${LCP_Offer_Rate_Bonus_Rate}    0.0007 per point
    Element Text Should Be    ${LCP_Offer_Rate_Commision_Rate}    0.00%
    Element Text Should Be    ${LCP_Offer_LPID}    1040270f-8e34-4ed2-890d-b1f0d9af58a1
    Comment    Element Text Should Be    ${LCP_Offer_Member_List_Count}    4
    Element Text Should Be    ${LCP_Offer_Tags}    buy, base, personalized offer, 25K
    Click element    ${LCP_Offer_Preview_Button}
    Wait Until Element Is Visible    ${LCP_Preview_Storefront_Name}    20s
    Input Text    ${LCP_Preview_Storefront_Name}    mileage-plan
    Wait Until Element Is Visible    ${LCP_Preview_Storefront_Type}    20s
    Select from list    ${LCP_Preview_Storefront_Type}    Standalone
    Click Element    ${LCP_Preview_form_previewButton}
    Sleep    10s
    Select window    url=https://storefront.points.com/mileage-plan/en-US/buy
    Wait Until Element Is Visible    ${bgt-offer-dropdown}    5s
    Select from list    ${bgt-offer-dropdown}    1000
    Capture Page Screenshot    filename=limit1_P-A.png
    Sleep    2s
    Set Window Size    375    1000
    Sleep    2s
    Capture Page Screenshot    filename=mobileHeader_P-A.png
    Comment    Click Element    ${StoreFront_Terms_CheckBox}
    Comment    Element Should be visible    ${StoreFront_Points_Logo_Legal}    Logo Image Exists    #verifies image
    ${width}    ${height}=    Get Window Size
    Log    ${width},${height}
    Sleep    2s
    : FOR    ${INDEX}    IN RANGE    1    2400    400    #For loop
    \    Scroll Page To Location    0    ${INDEX}    #To Scroll down
    \    Log    "Current Location", 0, ${INDEX}
    \    Sleep    2s
    \    Comment    Scroll Page To Location    0    -${INDEX}    #To scroll up
    \    Comment    Sleep    2s
    Capture Page Screenshot    filename=mobileLegal_P-A.png
    Maximize browser window
    : FOR    ${INDEX}    IN RANGE    1    1800    300    #For loop
    \    Scroll Page To Location    0    -${INDEX}    #To scroll up
    \    Comment    Sleep    2s
    Select from list    ${bgt-offer-dropdown}    3000
    Capture Page Screenshot    filename=limit2_P-A.png
    Sleep    2s
    Select from list    ${bgt-offer-dropdown}    5000
    Capture Page Screenshot    filename=limit3_P-A.png
    Sleep    2s
    Select from list    ${bgt-offer-dropdown}    10000
    Capture Page Screenshot    filename=limit4_P-A.png
    Sleep    2s
    Select from list    ${bgt-offer-dropdown}    15000
    Capture Page Screenshot    filename=limit5_P-A.png
    Sleep    2s
    Select from list    ${bgt-offer-dropdown}    20000
    Capture Page Screenshot    filename=limit6_P-A.png
    Sleep    2s
    Select from list    ${bgt-offer-dropdown}    30000
    Capture Page Screenshot    filename=limit7_P-A.png
    Sleep    2s
    Select from list    ${bgt-offer-dropdown}    40000
    Capture Page Screenshot    filename=limit8_P-A.png
    Sleep    2s
    Select from list    ${bgt-offer-dropdown}    50000
    Capture Page Screenshot    filename=limit9_P-A.png
    Sleep    2s
    Select from list    ${bgt-offer-dropdown}    60000
    Capture Page Screenshot    filename=limit10_P-A.png
    Sleep    2s
    Select from list    ${bgt-offer-dropdown}    1000
    Sleep    2s
    Element Text Should Be    ${introCopy}    Buying miles is the easy way to top up your account to get the award you want.
    Capture Page Screenshot    filename=Promo_Header_Banner_P-A.png
    Element Should Contain    ${legalCopy}    Miles are purchased from Points.com Inc. for a cost of $27.50 per 1,000 miles, plus a 7.5% Federal Excise Tax*, and GST/HST for Canadian residents. Miles are non-refundable and do not count toward MVP and MVP/Gold status. Offer is subject to change and all terms and conditions of the Mileage Plan Program apply.
    Input Text    ${StoreFront_Phone}    416-878-9090
    Sleep    2S
    Capture Page Screenshot    filename=Promo_Legal_P-A.png
    Sleep    2S

TC-1 Read Excel Data
    Open Browser and Login
    Enter First Name
    Enter Last Name
    Enter Member ID
    Enter Email
    sleep    2s
    Click Login
    sleep    15s
    Select Tier
    Verify Promo Copy

TC-1 Process A - PROD
    Open Browser and Login
    sleep    3s
    LCP Login
    sleep    5s
    Wait Until Element Is Visible    ${LCP_Offer_Tab_Production}    8s
    Click Element    ${LCP_Offer_Tab_Production}
    Sleep    5s
    Enter Offer Name
    Click element    ${LCP_Search_button}
    Wait Until Element Is Visible    ${LCP_Open_Offer}    40s
    Click element    ${LCP_Open_Offer}
    Wait Until Element Is Visible    ${LCP_Offer_Type}
    Scroll
    Verify Offer Config
    sleep    3s
    Comment    Click element    ${LCP_Offer_Edit_Button}
    Comment    ${LCP_Console_Offer_IntroCopy_Area}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    U    2
    Comment    Element should contain    ${LCP_Console_Offer_IntroCopy_AreaTextBox}    ${LCP_Console_Offer_IntroCopy_Area}
    Comment    Sleep    8s
    Comment    Element Should Contain    ${LCP_Console_Offer_IntroCopy_Area}    alt="A sign reading I wish you were here hangs outside the Snack Bar diner in Austin."
    Comment    sleep    2s
    Comment    Click element    ${LCP_Offer_Cancel_Button}
    Comment    sleep    5s
    Preview Offer
    sleep    2s
    Select Tier
    Verify Promo Copy
    Capture Page Screenshot    filename=Desktop_Header_P-A.png
    Scroll
    Capture Page Screenshot    filename=Desktop_Legal_P-A.png

*** Keywords ***
Scroll Page To Location
    [Arguments]    ${x_location}    ${y_location}
    Execute JavaScript    window.scrollTo(${xlocation},${y_location})

Open Browser and Login
    Comment    ${LCP_URL} =    Read Cell Value    Alaska_Data_Sheet.xlsx    Login_Creds    E    2
    Comment    Input Text    ${LCP_URLTextBox}    ${LCP_URL}
    Open Browser    https://admin.lcp.points.com/login    chrome
    Maximize Browser Window

Enter First Name
    ${firstname}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Login_Creds    A    2
    Input Text    ${firstnameTextBox}    ${firstname}

Enter Last Name
    ${lastname}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Login_Creds    B    2
    Input Text    ${lastnameTextBox}    ${lastname}

Enter Member ID
    ${memberid}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Login_Creds    C    2
    Input Text    ${MemnerIdTextBox}    ${memberid}

Enter Email
    ${email}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Login_Creds    D    2
    Input Text    ${EmailTextBox}    ${email}

Click Login
    Click Element    ${LoginButton}

Select Tier
    ${bgt-offer-dropdown}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    F    2
    Select from list    ${bgt-offer-dropdown_Textbox}    ${bgt-offer-dropdown}
    Sleep    2s
    Capture Page Screenshot    filename=limit1.png
    ${bgt-offer-dropdown}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    F    3
    Select from list    ${bgt-offer-dropdown_Textbox}    ${bgt-offer-dropdown}
    Sleep    2s
    Capture Page Screenshot    filename=limit2.png
    ${bgt-offer-dropdown}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    F    4
    Select from list    ${bgt-offer-dropdown_Textbox}    ${bgt-offer-dropdown}
    Sleep    2s
    Capture Page Screenshot    filename=limit3.png
    ${bgt-offer-dropdown}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    F    5
    Select from list    ${bgt-offer-dropdown_Textbox}    ${bgt-offer-dropdown}
    Sleep    2s
    Capture Page Screenshot    filename=limit4.png
    ${bgt-offer-dropdown}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    F    6
    Select from list    ${bgt-offer-dropdown_Textbox}    ${bgt-offer-dropdown}
    Sleep    2s
    Capture Page Screenshot    filename=limit5.png
    ${bgt-offer-dropdown}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    F    7
    Select from list    ${bgt-offer-dropdown_Textbox}    ${bgt-offer-dropdown}
    Sleep    2s
    Capture Page Screenshot    filename=limit6.png
    ${bgt-offer-dropdown}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    F    8
    Select from list    ${bgt-offer-dropdown_Textbox}    ${bgt-offer-dropdown}
    Sleep    2s
    Capture Page Screenshot    filename=limit7.png
    ${bgt-offer-dropdown}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    F    9
    Select from list    ${bgt-offer-dropdown_Textbox}    ${bgt-offer-dropdown}
    Sleep    2s
    Capture Page Screenshot    filename=limit8.png
    ${bgt-offer-dropdown}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    F    10
    Select from list    ${bgt-offer-dropdown_Textbox}    ${bgt-offer-dropdown}
    Sleep    2s
    Capture Page Screenshot    filename=limit9.png
    ${bgt-offer-dropdown}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    F    11
    Select from list    ${bgt-offer-dropdown_Textbox}    ${bgt-offer-dropdown}
    Sleep    2s
    Capture Page Screenshot    filename=limit10.png
    Set Window Size    375    1000
    Sleep    2s
    Capture Page Screenshot    filename=mobileHeader.png
    Comment    Click Element    ${StoreFront_Terms_CheckBox}
    Comment    Element Should be visible    ${StoreFront_Points_Logo_Legal}    Logo Image Exists    #verifies image
    ${width}    ${height}=    Get Window Size
    Log    ${width},${height}
    Sleep    2s
    : FOR    ${INDEX}    IN RANGE    1    2400    400    #For loop
    \    Scroll Page To Location    0    ${INDEX}    #To Scroll down
    \    Log    "Current Location", 0, ${INDEX}
    \    Sleep    2s
    Capture Page Screenshot    filename=mobileLegal.png
    Maximize browser window
    : FOR    ${INDEX}    IN RANGE    1    1800    300    #For loop
    \    Scroll Page To Location    0    -${INDEX}    #To scroll up

Verify Promo Copy
    ${introCopy}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    D    2
    Element Text Should Be    ${introCopyTextBox}    ${introCopy}
    ${legalCopy}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    E    2
    Element Should Contain    ${legalCopyTextBox}    ${legalCopy}

LCP Login
    ${LCP_email} =    Read Cell Value    Alaska_Data_Sheet.xlsx    Login_Creds    F    2
    Input Text    ${LCP_emailTextBox}    ${LCP_email}
    ${LCP_Password}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Login_Creds    G    2
    Input Text    ${LCP_PasswordTextBox}    ${LCP_Password}
    Click Element    ${LCP_Login_button}

Enter Offer Name
    ${LCP_OfferName} =    Read Cell Value    Alaska_Data_Sheet.xlsx    LCP_Offers    A    2
    Input Text    ${LCP_OfferName_FieldTextBox}    ${LCP_OfferName}

Verify Offer Config
    ${LCP_Offer_Type}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    O    2
    Element Text Should Be    ${LCP_Offer_TypeTextBox}    ${LCP_Offer_Type}
    ${LCP_Offer_Base_PIC}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    P    2
    Element Text Should Be    ${LCP_Offer_Base_PICTextBox}    ${LCP_Offer_Base_PIC}
    ${LCP_Offer_Start_Date}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    A    2
    Element Text Should Be    ${LCP_Offer_Start_DateTextBox}    ${LCP_Offer_Start_Date}
    ${LCP_Offer_End_Date}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    B    2
    Element Text Should Be    ${LCP_Offer_End_DateTextBox}    ${LCP_Offer_End_Date}
    ${LCP_Offer_Priority}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    C    2
    Element Text Should Be    ${LCP_Offer_PriorityTextBox}    ${LCP_Offer_Priority}
    ${LCP_Promo_Type}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    Q    2
    Element Text Should Be    ${LCP_Promo_TypeTextBox}    ${LCP_Promo_Type}
    ${LCP_Offer_Rate_Block_Size}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    G    2
    Element Text Should Be    ${LCP_Offer_Rate_Block_SizeTextBox}    ${LCP_Offer_Rate_Block_Size}
    ${LCP_Offer_Rate_Effective}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    K    2
    Element Text Should Be    ${LCP_Offer_Rate_EffectiveTextBox}    ${LCP_Offer_Rate_Effective}
    ${LCP_Offer_Rate_Wholesale}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    L    2
    Element Text Should Be    ${LCP_Offer_Rate_WholesaleTextBox}    ${LCP_Offer_Rate_Wholesale}
    ${LCP_Offer_Rate_Bonus_Rate}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    M    2
    Element Text Should Be    ${LCP_Offer_Rate_Bonus_RateTextBox}    ${LCP_Offer_Rate_Bonus_Rate}
    ${LCP_Offer_Rate_Commision_Rate}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    N    2
    Element Text Should Be    ${LCP_Offer_Rate_Commision_RateTextBox}    ${LCP_Offer_Rate_Commision_Rate}
    ${LCP_Offer_LPID}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    R    2
    Element Text Should Be    ${LCP_Offer_LPIDTextBox}    ${LCP_Offer_LPID}
    Comment    ${LCP_Offer_Member_List_Count}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    S    2
    Comment    Element Text Should Be    ${LCP_Offer_Member_List_CountTextBox}    ${LCP_Offer_Member_List_Count}
    ${LCP_Offer_Tags}=    Read Cell Value    Alaska_Data_Sheet.xlsx    Promo_Details    T    2
    Element Text Should Be    ${LCP_Offer_TagsTextBox}    ${LCP_Offer_Tags}

Scroll
    Capture Page Screenshot    filename=Header_P-A.png
    ${width}    ${height}=    Get Window Size
    Log    ${width},${height}
    Sleep    2s
    : FOR    ${INDEX}    IN RANGE    1    2400    400    #For loop
    \    Scroll Page To Location    0    ${INDEX}    #To Scroll down
    \    Log    "Current Location", 0, ${INDEX}
    \    Capture Page Screenshot
    \    Sleep    2s
    \    Comment    Scroll Page To Location    0    -${INDEX}    #To scroll up
    \    Comment    Sleep    2s
    Capture Page Screenshot    filename=Footer_P-A.png
    : FOR    ${INDEX}    IN RANGE    1    1800    300    #For loop
    \    Scroll Page To Location    0    -${INDEX}    #To scroll up
    \    Comment    Sleep    2s

Mobile Version
    Set Window Size    375    1000
    Sleep    2s
    Capture Page Screenshot    filename=mobileHeader_P-A.png
    ${width}    ${height}=    Get Window Size
    Log    ${width},${height}
    Sleep    2s
    : FOR    ${INDEX}    IN RANGE    1    2400    400    #For loop
    \    Scroll Page To Location    0    ${INDEX}    #To Scroll down
    \    Log    "Current Location", 0, ${INDEX}
    \    Sleep    2s
    \    Comment    Scroll Page To Location    0    -${INDEX}    #To scroll up
    \    Comment    Sleep    2s
    Capture Page Screenshot    filename=mobileLegal_P-A.png
    Maximize browser window
    : FOR    ${INDEX}    IN RANGE    1    1800    300    #For loop
    \    Scroll Page To Location    0    -${INDEX}    #To scroll up
    \    Comment    Sleep    2s

Preview Offer
    Click element    ${LCP_Offer_Preview_Button}
    Wait Until Element Is Visible    ${LCP_Preview_Storefront_Name}    20s
    Input Text    ${LCP_Preview_Storefront_Name}    mileage-plan
    Wait Until Element Is Visible    ${LCP_Preview_Storefront_Type}    20s
    Select from list    ${LCP_Preview_Storefront_Type}    Standalone
    Click Element    ${LCP_Preview_form_previewButton}
    Sleep    10s
    Select window    url=https://storefront.points.com/mileage-plan/en-US/buy
    Wait Until Element Is Visible    ${bgt-offer-dropdown}    5s
