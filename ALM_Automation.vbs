'UFT’s new UI Automation add-in

With UIAWindow("HPE Application Lifecycle").UIAWindow("MainForm")
                'Set user name
                .UIATab("m_loginTabControl").UIAEdit("m_user").SetValue “user name”
                'Set password, to keep a password in encrypted form, use the Password Encoder tool.
                .UIATab("m_loginTabControl").UIAEdit("m_password").SetSecure “password”
                'Click the Authenticate button
                .UIATab("m_loginTabControl").UIAButton("Authenticate").Click
                'Set a domain
                .UIATab("m_loginTabControl").UIAComboBox("m_domains").Select “domain”
                'Set a project
                .UIATab("m_loginTabControl").UIAComboBox("m_projects").Select “project”
                'Login
                .UIATab("m_loginTabControl").UIAButton("Login").Click
End With
wait (5)

'Select the Defects item in the left side list
UIAWindow("HP Application Lifecycle").UIAWindow("MainForm").UIAList("m_modulesNavigator").ClickItem "Defects"

'Open a new defect

UIAWindow("HP Application Lifecycle").UIAWindow("MainForm").UIAButton("&New Defect...").Click

'Set a summary of the defect

UIAWindow("HP Application Lifecycle").UIAWindow("EntityForm").UIATab("m_paneTabControl").UIAComboBox("BG_SUMMARY").Type “summary”

With UIAWindow("HP Application Lifecycle").UIAWindow("EntityForm").UIATab("m_mainTabControl")

    'Set a severity of the defect

    .UIATab("m_tab").UIATab("m_tabUpper").UIAComboBox("BG_SEVERITY").Type “severity”

    'Fill a description of the defect (steps to reproduce etc)

    .UIATab("m_tab").UIATab("m_tabLower").UIAObject("BG_DESCRIPTION").SetValue DataTable("Description", dtGlobalSheet)

End With

'Submit the defect

UIAWindow("HP Application Lifecycle").UIAWindow("EntityForm").UIAButton("Submit").Click

'https://community.hpe.com/t5/All-About-the-Apps/4-Steps-to-ALM-automation-with-automated-testing/ba-p/6952164#.WQNXbPmGPIV
