# MS_Access_Pasword_Vault
Password Vault for Microsoft Access

Requires two tables:
  tblMaster: CREATE TABLE tblMaster (id INT NOT NULL AUTO_INCREMENT, Master TEXT (255));
  tblVault: CREATE TABLE tblVault (id INT NOT NULL AUTO_INCREMENT, Service TEXT (255), Password TEXT (255), Username TEXT (255));
  
Requires six forms:
  frmLogin:
    lblMasterPass: Label ("Master Password:")
    txtMasterPass: Textbox (255)
    btnSubmit: CommandButton
    Popup: True
    Modal: True
    Control Source: None
    Subform: None
    
  frmSetMasterPassword:
    lblMasterPass: Label ("Master Password:")
    txtMasterPass: Textbox (255), Format=Password
    lblConfMasterPass: Label ("Confirm Master Password:")
    txtConfMasterPass: Textbox (255), Format=Password
    btnSubmit: CommandButton
    Popup: True
    Modal: True
    Control Source: None
    Subform: None
   
  frmSetServicePassword:
    lblService: Label ("Service:")
    txtService: Textbox (255)
    lblUser: Label ("Username:")
    txtUser: Textbox (255)
    lblPass: Label ("Password:")
    txtPass: Textbox (255)
    lblConfPass: Label ("Confirm Password:")
    txtConfPass: Textbox (255)
    btnSubmit: CommandButton
    btnRandomPW: CommandButton
    btnDelete: CommandButton
    Popup: True
    Modal: False
    Control Source: None
    Subform: None
  
  frmStartup:
    (Empty Form)
    Popup: False
    Modal: False
    Control Source: None
    Subform: None
  
  frmVault:
    btnAdd: CommandButton
    Popup: True
    Modal: False
    Control Source: None
    Subform: fsubVault
    
  fsubVault:
    lblService: Label ("Service:")
    txtService: Textbox (255)
    lblUser: Label ("Username:")
    txtUser: Textbox (255)
    lblPass: Label ("Password:")
    txtPass: Textbox (255)
    Popup: True
    Modal: False
    Control Source: tblVault
    Subform: None
    
