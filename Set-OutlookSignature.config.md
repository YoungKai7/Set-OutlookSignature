```xml
 <config>
    <main>
        <UserName>$env:username</UserName>
        <SigSource>.\signature_data</SigSource><!-- path to folder containing the signature template(s) and Excel user profiles -->
        <UserSource>Excel</UserSource><!-- "Excel" or "AD" -->
        <UserSourceFile>$SigSource\UserDirectory.xlsx</UserSourceFile><!-- only for Excel user source -->
    </main>
    <profile>
        <!--
            Values can reference user profile as $($User) if needed. Ex. $($User.company)

            Different templates can be used based on Company and/or Department.
            Save template file into subfolders and name the subfolders according to the Company
            or Department names.

            Template priority: Department > Company > Default
        -->
        <Company>$env:userdomain</Company>
        <Department>$($User.Department)</Department>
        <regKey comment="Windows registry key where update history will be saved to"
            >$env:userdomain</regKey>
        <signatureName comment="Outlook signature name and user signation file names"
            >$env:userdomain (AUTO-SIG)</signatureName>
        <TemplateName comment="name of Word doc signature template"
            >Unified-Signature.docx</TemplateName>
        <ForceSignatureNew comment="Set as default signature for new messages. 0 = Not Set, 1 = Set"
            >1</ForceSignatureNew>
        <ForceSignatureReplyForward comment="Set as default signature for reply and forward messages. 0 = Not Set, 1 = Set"
            >1</ForceSignatureReplyForward>
    </profile>
    <customProfile>
        <!--
            Add custom nodes here to use in signature template.
            Values can be static or it can reference user profile as $($User) if needed. Ex. $($User.Company)
            
            Name of each node here should match the user property name.
            i.e. User Property name in AD, or column heading in Excel.
            The values used from default config matches standard AD user properties.
        -->
        <CompanyName>$($User.Company)</CompanyName>
        <FirstName>$($User.FirstName)</FirstName>
        <LastName>$($User.LastName)</LastName>
        <FirstLastName>$($User.FirstName) $($User.LastName)</FirstLastName>
        <FullName>$($User.DisplayName)</FullName>
        <Title>$($User.Title)</Title>
        <Telephone>$($User.TelephoneNumber)</Telephone>
        <Mobile>$($User.Mobile)</Mobile>
        <Email>$($User.Mail)</Email>
        <DepartmentName>$($User.Department)</DepartmentName>
    </customProfile>
</config>
```
