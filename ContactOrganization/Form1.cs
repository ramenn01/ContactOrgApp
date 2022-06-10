//Make dictionary contain 1 email and a list of contacts. Every email gets it's own dictionary entry, but if ANY email is shared, the contact list is combined (they become the same list, changing one would change the other)            

using Microsoft.Exchange.WebServices.Data;
using System.Reflection;



//make sure to combine if they have the same numbers. 






//The address doesn't display on the contact, although the information is there if you go in to edit it. 

//Can't add item attachments. 


//Note body has no duplicate protection.



//Below are properties that are found in a contact, but I haven't found yet 
/*
 * CAN'T FIND JOURNAL
 * CAN'T FIND ISDN
 * CAN'T FIND ICON IN PROPERTIES OF A CONTACT (CAN FIND ICON INDEX)
 * CAN'T FIND COMPUTER NETWORK NAME IN PROPERTIES OF A CONTACT
 * CAN'T FIND CONTACTS IN PROPERTIES OF A CONTACT
 * CAN'T FIND FOLLOW UP FLAG IN PROPERTIES OF A CONTACT (MIGHT BE COVERED BY OVERALL FLAG COMBINING)
 * CAN'T FIND PERSONAL HOME PAGE
 * CAN'T FIND GENDER, HOBBIES, LANGUAGE, REFFERED BY
             
*/








/*    THIS SECTION HOLDS ALL THE NON-SINGLE PROPERTY VALUES I NEED TO CODE INDIVIDUALLY, PROBABLY BECAUSE THEY ARE GROUPS OF DATA (ARRAYS, LISTS)
            *    Attachments      (This one needs item attachments, not file attachments. 
            *    ExtendedProperties
            *    IconIndex
            *    InstanceKey
            *    InternetMessageHeaders
            *    ManagerMailbox
            *    MimeContent
            *    NormalizedBody
            *    MSExchangeCertificate
            *    ParentFolderId
            *    Schema
            *    Service
            *    
            *    
            *    
            *    
            *    THE MOST IMPORTANT ONES ARE EXTENDED PROPERTIES, AND MIMECONTENT, BECAUSE I DON'T KNOW WHAT THEY ARE
            *    Extended properties are custom properties made outside outlook. 

           */

/*     THIS SECTION HOLDS ALL THE PROPERTY VALUES WITHOUT A SET VALUE (PROBABLY DON'T NEED THEM. STILL LOOK INTO IT)
 *     DateTimeCreated
 *     ConversationId
 *     Contactsource
 *     AllowedResponseActions
 *     Alias
 *     IsUnmodified
 *     LastModifiedName
 *     LastModifiedTime
 *     DirectoryId
 *     DirectoryPhoto
 *     DirectoryReports
 *     DisplayCc
 *     DisplayTo

 */



namespace ContactOrganization
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            textBoxEmail.Text = "trevor.magruder@albionnet.com";
            textBoxPassword.Text = "B0bth3Guy";
        }




        private Contact combineContacts(Contact origional, Contact second)
        {
            List<string> subOverflow = new List<string>();
            Contact[] contactList = new Contact[2] { origional, second };


            //UPDATE NOTE BODY


            //StringReader sr = new StringReader(origional.Body.Text);
            StringReader sr2 = new StringReader(second.Body.Text);
            origional.Body.Text += sr2.ReadToEnd(); 



            /*bool orgIsHtml = false;
            bool secIsHtml = false;
            string str = null;
            string str2 = null;
            string cbody = "";
            string sbody = "";

            try
            {
                StringReader sr = new StringReader(origional.Body.Text);
                str = sr.ReadLine();
                if (str.Contains("<html>"))
                {
                    orgIsHtml = true;
                    str += Environment.NewLine + sr.ReadToEnd();
                    cbody = str.Split("<body>")[1];
                    cbody = cbody.Split("</body>")[0]; ;
                }
                else
                {
                    cbody += str;
                    while (1 == 1)
                    {
                        str = sr.ReadLine();
                        if (str == null)
                        {
                            break;
                        }
                        else
                        {
                            cbody += Environment.NewLine + str;
                        }

                    }
                }
            }
            catch (ArgumentNullException)
            {

            }
            try
            {
                StringReader sr2 = new StringReader(second.Body.Text);
                str2 = sr2.ReadLine();
                if (str2.Contains("<html>"))
                {
                    str2 += Environment.NewLine + sr2.ReadToEnd();
                    sbody = str2.Split("<body>")[1];
                    sbody = sbody.Split("</body>")[0];
                    secIsHtml = true;
                }
                else
                {
                    sbody += str2;
                    while (1 == 1)
                    {
                        str2 = sr2.ReadLine();
                        if (str2 == null)
                        {
                            break;
                        }
                        else
                        {
                            sbody += str2;
                        }

                    }
                }
            }
            catch (ArgumentNullException)
            { }

            if (orgIsHtml == true)
            {
                if (secIsHtml == true)
                {
                    str = str.Split("<body>")[0] + "<body>" + cbody + sbody + "</body>" + str.Split("</body>")[1];
                    origional.Body.Text = str;
                }
                else
                {
                    origional.Body.Text = str.Split("<body>")[0] + "<body>" + cbody + Environment.NewLine + "<div>" + sbody + "</div>" + Environment.NewLine + "</body>" + str.Split("</body>")[1];
                }

            }
            else
            {
                if (secIsHtml == true)
                {
                    origional.Body.Text = str2.Split("<body>")[0] + "<body>" + Environment.NewLine + "<div>" + cbody + "</div>" + Environment.NewLine + sbody + "</body>" + str.Split("</body>")[1];
                }
                else
                {
                    origional.Body.Text = cbody + Environment.NewLine + sbody;
                }
            }
            */



            //UPDATE EMAIL ADDRESSES

            List<EmailAddressKey> emailKeys = new List<EmailAddressKey>();
            emailKeys.Add(EmailAddressKey.EmailAddress1);
            emailKeys.Add(EmailAddressKey.EmailAddress2);
            emailKeys.Add(EmailAddressKey.EmailAddress3);


            List<string> usedAddresses = new List<string>();
            List<string> emails = new List<string>();
            List<string> addedAddresses = new List<string>();

            for (int i = 0; i < 3; i++)
            {
                if (origional.EmailAddresses.Contains(emailKeys[i]))
                {
                    if (origional.EmailAddresses[emailKeys[i]].Address != null && usedAddresses.Contains(origional.EmailAddresses[emailKeys[i]].Address) == false)
                    {
                        usedAddresses.Add(origional.EmailAddresses[emailKeys[i]].Address);
                        addedAddresses.Add(origional.EmailAddresses[emailKeys[i]].Address);
                    }
                }
            }
            //Example here
            for (int i = 0; i < 3; i++)
            {
                if (second.EmailAddresses.Contains(emailKeys[i]))
                {
                    if (second.EmailAddresses[emailKeys[i]].Address != null && usedAddresses.Contains(second.EmailAddresses[emailKeys[i]].Address) == false)
                    {
                        emails.Add(second.EmailAddresses[emailKeys[i]].Address);
                        usedAddresses.Add(second.EmailAddresses[emailKeys[i]].Address);
                    }
                }
            }
            foreach (string email in emails)
            {
                foreach (EmailAddressKey key in emailKeys)
                {
                    if (origional.EmailAddresses.Contains(key))
                    {
                        if (origional.EmailAddresses[key].Address == null)
                        {
                            origional.EmailAddresses[key] = email;
                            addedAddresses.Add(origional.EmailAddresses[key].Address);
                            break;
                        }
                    }
                    else
                    {
                        origional.EmailAddresses[key] = email;
                        addedAddresses.Add(origional.EmailAddresses[key].Address);
                        break;
                    }
                }
            }
            foreach (EmailAddress email in emails)
            {
                if (addedAddresses.Contains(email.Address) == false)
                {
                    subOverflow.Add(origional.DisplayName + " Email Address: " + email.Address);
                }
            }
            //Attachments

            foreach (Attachment item in second.Attachments)
            {
                if (item is FileAttachment)
                {
                    if (origional.Attachments.Where(p => p.Name == item.Name).FirstOrDefault() == null)
                    {
                        FileAttachment fileAttachment = item as FileAttachment;
                        if (fileAttachment.Id != null)
                        {
                            fileAttachment.Load();
                        }
                        origional.Attachments.AddFileAttachment(fileAttachment.Name, fileAttachment.Content);

                    }
                }
            }



            //PhoneNumbers

            //HomePhone
            PhoneNumberKey[] homePhoneKeys = new PhoneNumberKey[2] { PhoneNumberKey.HomePhone, PhoneNumberKey.HomePhone2 };
            PhoneNumberKey[] busPhoneKeys = new PhoneNumberKey[2] { PhoneNumberKey.BusinessPhone, PhoneNumberKey.BusinessPhone2 };

            List<string> homePhoneNums = new List<string>();
            /*List<PhoneNumberKey> homePhoneKeys = new List<PhoneNumberKey>();
            homePhoneKeys.Add(PhoneNumberKey.HomePhone);
            homePhoneKeys.Add(PhoneNumberKey.HomePhone2);
            List<PhoneNumberKey> busPhoneKeys = new List<PhoneNumberKey>();
            busPhoneKeys.Add(PhoneNumberKey.BusinessPhone);
            busPhoneKeys.Add(PhoneNumberKey.BusinessPhone2);*/

            foreach (PhoneNumberKey phoneKey in homePhoneKeys)
            {
                foreach (Contact con in contactList)
                {
                    if (con.PhoneNumbers.Contains(phoneKey))
                    {
                        if (con.PhoneNumbers[phoneKey] != null)
                        {
                            if (homePhoneNums.Contains(con.PhoneNumbers[phoneKey]) == false)
                            {
                                homePhoneNums.Add(con.PhoneNumbers[phoneKey]);
                            }

                        }
                    }
                }
            }
            for (int i = 0; i < homePhoneNums.Count; i++)
            {
                if (i > 1)
                {
                    subOverflow.Add(origional.DisplayName + " Home Phone Number : " + homePhoneNums[i]);
                }
                else
                {
                    if (i == 0)
                    {
                        origional.PhoneNumbers[PhoneNumberKey.HomePhone] = homePhoneNums[i];
                    }
                    if (i == 1)
                    {
                        origional.PhoneNumbers[PhoneNumberKey.HomePhone2] = homePhoneNums[i];
                    }
                }
            }

            //BusinessPhone

            List<string> busPhoneNums = new List<string>();

            foreach (PhoneNumberKey phoneKey in busPhoneKeys)
            {
                foreach (Contact con in contactList)
                {
                    if (con.PhoneNumbers.Contains(phoneKey))
                    {
                        if (con.PhoneNumbers[phoneKey] != null)
                        {
                            if (busPhoneNums.Contains(con.PhoneNumbers[phoneKey]) == false)
                            {
                                busPhoneNums.Add(con.PhoneNumbers[phoneKey]);
                            }
                        }
                    }
                }
            }
            for (int i = 0; i < busPhoneNums.Count; i++)
            {
                if (i > 1)
                {
                    subOverflow.Add(origional.DisplayName + " Business Phone Number : " + busPhoneNums[i]);
                }
                else
                {
                    if (i == 0)
                    {
                        origional.PhoneNumbers[PhoneNumberKey.BusinessPhone] = busPhoneNums[i];
                    }
                    if (i == 1)
                    {
                        origional.PhoneNumbers[PhoneNumberKey.BusinessPhone2] = busPhoneNums[i];
                    }
                }
            }

            //MiscPhone
            PhoneNumberKey[] NumberTypes = new PhoneNumberKey[]
            {
                PhoneNumberKey.TtyTddPhone,
                PhoneNumberKey.AssistantPhone,
                PhoneNumberKey.BusinessFax,
                PhoneNumberKey.Callback,
                PhoneNumberKey.CarPhone,
                PhoneNumberKey.CompanyMainPhone,
                PhoneNumberKey.HomeFax,
                PhoneNumberKey.Isdn,
                PhoneNumberKey.MobilePhone,
                PhoneNumberKey.OtherFax,
                PhoneNumberKey.OtherTelephone,
                PhoneNumberKey.Pager,
                PhoneNumberKey.PrimaryPhone,
                PhoneNumberKey.RadioPhone,
                PhoneNumberKey.TtyTddPhone,
                PhoneNumberKey.Telex
            };

            foreach (PhoneNumberKey type in NumberTypes)
            {
                if (origional.PhoneNumbers.Contains(type))
                {
                    if (origional.PhoneNumbers[type] != null && second.PhoneNumbers.Contains(type))
                    {
                        if (origional.PhoneNumbers[type] != second.PhoneNumbers[type] && second.PhoneNumbers[type] != null)
                        {
                            subOverflow.Add(origional.DisplayName + " " + type.ToString() + " : " + second.PhoneNumbers[type]);
                        }
                    }
                    else if (second.PhoneNumbers.Contains(type))
                    {
                        origional.PhoneNumbers[type] = second.PhoneNumbers[type];
                    }
                }
                else
                {
                    if (second.PhoneNumbers.Contains(type))
                    {
                        if (second.PhoneNumbers[type] != null)
                        {
                            origional.PhoneNumbers[type] = second.PhoneNumbers[type];
                        }
                    }
                }
            }


            //Physical Addresses
            List<PhysicalAddressKey> physicalAddresses = new List<PhysicalAddressKey>();
            physicalAddresses.Add(PhysicalAddressKey.Business);
            physicalAddresses.Add(PhysicalAddressKey.Home);
            physicalAddresses.Add(PhysicalAddressKey.Other);

            foreach (PhysicalAddressKey type in physicalAddresses)
            {
                if (origional.PhysicalAddresses.Contains(type))
                {
                    if (second.PhysicalAddresses.Contains(type))
                    {
                        if (origional.PhysicalAddresses[type].PostalCode == second.PhysicalAddresses[type].PostalCode && origional.PhysicalAddresses[type].State == second.PhysicalAddresses[type].State && origional.PhysicalAddresses[type].City == second.PhysicalAddresses[type].City && origional.PhysicalAddresses[type].CountryOrRegion == second.PhysicalAddresses[type].CountryOrRegion && origional.PhysicalAddresses[type].Street == second.PhysicalAddresses[type].Street)
                        {

                        }
                        else
                        {
                            if (origional.PhysicalAddresses[type].PostalCode == null && origional.PhysicalAddresses[type].State == null && origional.PhysicalAddresses[type].City == null && origional.PhysicalAddresses[type].CountryOrRegion == null && origional.PhysicalAddresses[type].Street == null)
                            {
                                origional.PhysicalAddresses[type] = second.PhysicalAddresses[type];
                                origional.PhysicalAddresses[type].City = second.PhysicalAddresses[type].City;
                                origional.PhysicalAddresses[type].PostalCode = second.PhysicalAddresses[type].PostalCode;
                                origional.PhysicalAddresses[type].CountryOrRegion = second.PhysicalAddresses[type].CountryOrRegion;
                                origional.PhysicalAddresses[type].State = second.PhysicalAddresses[type].State;
                                origional.PhysicalAddresses[type].Street = second.PhysicalAddresses[type].Street;
                            }
                            else
                            {
                                if (second.PhysicalAddresses[type].PostalCode == null && second.PhysicalAddresses[type].State == null && second.PhysicalAddresses[type].City == null && second.PhysicalAddresses[type].CountryOrRegion == null && second.PhysicalAddresses[type].Street == null)
                                {

                                }
                                else
                                {
                                    string overflowed = type.ToString() + " Address : ";
                                    //added "if null" to make this look neater, takes longer I guess
                                    //example

                                    if (second.PhysicalAddresses[type].Street != null)
                                    {
                                        overflowed += second.PhysicalAddresses[type].Street + ", ";
                                    }
                                    if (second.PhysicalAddresses[type].City != null)
                                    {
                                        overflowed += second.PhysicalAddresses[type].City + ", ";
                                    }
                                    if (second.PhysicalAddresses[type].State != null)
                                    {
                                        overflowed += second.PhysicalAddresses[type].State + ", ";
                                    }
                                    if (second.PhysicalAddresses[type].PostalCode != null)
                                    {
                                        overflowed += second.PhysicalAddresses[type].PostalCode + ", ";
                                    }
                                    if (second.PhysicalAddresses[type].CountryOrRegion != null)
                                    {
                                        overflowed += second.PhysicalAddresses[type].CountryOrRegion + ", ";
                                    }
                                    subOverflow.Add(overflowed);
                                }
                            }
                        }
                    }

                }

                else
                {
                    if (second.PhysicalAddresses.Contains(type))
                    {
                        origional.PhysicalAddresses[type] = second.PhysicalAddresses[type];
                        origional.PhysicalAddresses[type].City = second.PhysicalAddresses[type].City;
                        origional.PhysicalAddresses[type].PostalCode = second.PhysicalAddresses[type].PostalCode;
                        origional.PhysicalAddresses[type].CountryOrRegion = second.PhysicalAddresses[type].CountryOrRegion;
                        origional.PhysicalAddresses[type].State = second.PhysicalAddresses[type].State;
                        origional.PhysicalAddresses[type].Street = second.PhysicalAddresses[type].Street;
                    }
                }
            }


            //IMAddresses

            List<string> orgKeys = new List<string>();
            List<string> secKeys = new List<string>();

            if (origional.ImAddresses.Contains(ImAddressKey.ImAddress1))
            {
                if (origional.ImAddresses[ImAddressKey.ImAddress1] != null)
                {
                    orgKeys.Add(origional.ImAddresses[ImAddressKey.ImAddress1]);
                }
            }
            if (origional.ImAddresses.Contains(ImAddressKey.ImAddress2))
            {
                if (origional.ImAddresses[ImAddressKey.ImAddress2] != null && orgKeys.Contains(origional.ImAddresses[ImAddressKey.ImAddress2]) == false)
                {
                    orgKeys.Add(origional.ImAddresses[ImAddressKey.ImAddress2]);
                }
            }
            if (origional.ImAddresses.Contains(ImAddressKey.ImAddress3))
            {
                if (origional.ImAddresses[ImAddressKey.ImAddress3] != null && orgKeys.Contains(origional.ImAddresses[ImAddressKey.ImAddress3]) == false)
                {
                    orgKeys.Add(origional.ImAddresses[ImAddressKey.ImAddress3]);
                }
            }
            if (second.ImAddresses.Contains(ImAddressKey.ImAddress1))
            {
                if (second.ImAddresses[ImAddressKey.ImAddress1] != null && orgKeys.Contains(second.ImAddresses[ImAddressKey.ImAddress1]) == false)
                {
                    secKeys.Add(second.ImAddresses[ImAddressKey.ImAddress1]);
                }
            }
            if (second.ImAddresses.Contains(ImAddressKey.ImAddress2))
            {
                if (second.ImAddresses[ImAddressKey.ImAddress2] != null && secKeys.Contains(second.ImAddresses[ImAddressKey.ImAddress2]) == false && orgKeys.Contains(second.ImAddresses[ImAddressKey.ImAddress2]) == false)
                {
                    secKeys.Add(second.ImAddresses[ImAddressKey.ImAddress2]);
                }
            }
            if (second.ImAddresses.Contains(ImAddressKey.ImAddress3))
            {
                if (second.ImAddresses[ImAddressKey.ImAddress3] != null && secKeys.Contains(second.ImAddresses[ImAddressKey.ImAddress3]) == false && orgKeys.Contains(second.ImAddresses[ImAddressKey.ImAddress3]) == false)
                {
                    secKeys.Add(second.ImAddresses[ImAddressKey.ImAddress3]);
                }
            }
            ImAddressKey[] keys = new ImAddressKey[] { ImAddressKey.ImAddress1, ImAddressKey.ImAddress2, ImAddressKey.ImAddress3 };

            for (int i = 0; i < orgKeys.Count + secKeys.Count; i++)
            {
                if (i > 0)
                {
                    if (i + 1 > orgKeys.Count)
                    {
                        subOverflow.Add(secKeys[i - orgKeys.Count]);
                    }
                    else
                    {
                        subOverflow.Add(orgKeys[i]);
                    }
                    //adding everything that isn't the one set to origional slot 1 to overflow addresses because I can't see them shown in outlook. Just a backup, delete later. 
                }
                if (i <= 2)

                {
                    if (i + 1 > orgKeys.Count)
                    {
                        if (i + 1 - orgKeys.Count > secKeys.Count)
                        {
                            string stopbecausedanger = "this should never trigger";
                        }
                        else
                        {
                            origional.ImAddresses[keys[i]] = secKeys[i - orgKeys.Count];
                        }
                    }
                    else
                    {
                        origional.ImAddresses[keys[i]] = orgKeys[i];
                    }
                }

            }


            /* CAN'T FIND JOURNAL
             * CAN'T FIND ISDN
             * CAN'T FIND ICON IN PROPERTIES OF A CONTACT (CAN FIND ICON INDEX)
             * CAN'T FIND COMPUTER NETWORK NAME IN PROPERTIES OF A CONTACT
             * CAN'T FIND CONTACTS IN PROPERTIES OF A CONTACT
             * CAN'T FIND FOLLOW UP FLAG IN PROPERTIES OF A CONTACT (MIGHT BE COVERED BY OVERALL FLAG COMBINING)
             * CAN'T FIND PERSONAL HOME PAGE
             * CAN'T FIND GENDER, HOBBIES, LANGUAGE, REFFERED BY
             */


            List<PropertyInfo> contactProps = new List<PropertyInfo>();

            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.Department)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.CompanyName)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.BusinessHomePage)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.JobTitle)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.OfficeLocation)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.ArchiveTag)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.AssistantName)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.Birthday)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.Culture)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.FileAsMapping))); //Default SurnameCommaGivenName
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.Generation)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.GivenName)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.Importance))); //Default Normal
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.InReplyTo)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.IsReminderSet))); //Default false
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.ItemClass))); //Default IPM.Contact
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.Manager)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.MiddleName)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.Mileage)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.MimeContent)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.NickName)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.Profession)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.PolicyTag)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.PostalAddressIndex)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.ReminderDueBy)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.ReminderMinutesBeforeStart))); //Default 0
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.SpouseName)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.Surname)));
            contactProps.Add(typeof(Contact).GetProperty(nameof(origional.WeddingAnniversary)));





            /*    THIS SECTION HOLDS ALL THE NON-SINGLE PROPERTY VALUES I NEED TO CODE INDIVIDUALLY
             *    Attachments      (This one needs item attachments, not file attachments. 
             *    ExtendedProperties
             *    IconIndex
             *    InstanceKey
             *    InternetMessageHeaders
             *    ManagerMailbox
             *    MimeContent
             *    NormalizedBody
             *    MSExchangeCertificate
             *    ParentFolderId
             *    Schema
             *    Service
             *    
             *    
             *    
             *    
             *    THE MOST IMPORTANT ONES ARE EXTENDED PROPERTIES, AND MIMECONTENT AND NORMALIZED BODY, WHICH I DON'T KNOW WHAT THEY ARE

            */








            /*     THIS SECTION HOLDS ALL THE PROPERTY VALUES WITHOUT A SET VALUE (PROBABLY DON'T NEED THEM)
             *     DateTimeCreated
             *     ConversationId
             *     Contactsource
             *     AllowedResponseActions
             *     Alias
             *     IsUnmodified
             *     LastModifiedName
             *     LastModifiedTime
             *     DirectoryId
             *     DirectoryPhoto
             *     DirectoryReports
             *     DisplayCc
             *     DisplayTo
            
             */

            /*    THIS SECTION HOLDS ALL THE PROPERTY VALUES I PROBABLY STILL NEED TO CHANGE
             *    Attachments (item attachments is what I need, files already done)
             *    ExtendedProperties
             *    ManagerMailbox
             *    */


            foreach (PropertyInfo p in contactProps)
            {
                try
                {
                    if (p.GetValue(second) != null)
                    {
                        if (p.GetValue(origional) != null)
                        {
                            if (p.GetValue(origional).ToString() != p.GetValue(second).ToString())
                            {
                                subOverflow.Add(origional.DisplayName + " " + p.Name + " : " + p.GetValue(second).ToString());
                            }
                        }
                        else
                        {
                            p.SetValue(origional, p.GetValue(second));
                        }
                    }
                }
                catch (TargetInvocationException)
                {
                    try
                    {
                        if (p.GetValue(second) != null)
                        {
                            p.SetValue(origional, p.GetValue(second));
                        }
                    }
                    catch (TargetInvocationException)
                    {

                    }
                }

            }

            //Children

            foreach (var child in second.Children)
            {
                if (origional.Children.Contains(child) == false)
                {
                    origional.Children.Add(child);
                }
            }


            //categories

            foreach (string stri in second.Categories)
            {
                if (origional.Categories.Contains(stri) == false)
                {
                    origional.Categories.Add(stri);
                }
            }

            //Company

            foreach (string company in second.Companies)
            {
                if (origional.Companies.Contains(company) == false)
                {
                    origional.Companies.Add(company);
                }
            }

            //Sensitivity

            if (origional.Sensitivity == Sensitivity.Normal)
            {
                origional.Sensitivity = second.Sensitivity;
            }
            else
            {
                if (second.Sensitivity != Sensitivity.Normal)
                {
                    subOverflow.Add("Sensitivity: " + second.Sensitivity);
                }
            }

            //Flag
            if (origional.Flag.FlagStatus == ItemFlagStatus.NotFlagged)
            {
                origional.Flag.FlagStatus = second.Flag.FlagStatus;
            }
            else
            {
                if (second.Flag.FlagStatus != ItemFlagStatus.NotFlagged)
                {
                    subOverflow.Add("Flag: " + second.Flag.FlagStatus);
                }
            }

            //Display Name, Initials, Subject, Given Name, Surname, and FileName

            if (second.DisplayName.Contains(origional.DisplayName))
            {
                origional.DisplayName = second.DisplayName;
                origional.FileAs = second.FileAs;
                origional.Initials = second.Initials;
                origional.Subject = second.Subject;
                origional.GivenName = second.GivenName;
                origional.Surname = second.Surname;
            }



            //Adding overflow information to note body. 

            origional.Body.Text += Environment.NewLine + Environment.NewLine + "Information Lost From The Combined Contacts Below:";
            foreach (string overflowedInfo in subOverflow)
            {
                origional.Body.Text += Environment.NewLine + overflowedInfo;
            }
            origional.Body.Text += Environment.NewLine + "End Of Lost Info";

            /*string st = null;


            try
            {
                StringReader strRdr = new StringReader(origional.Body.Text);
                st = strRdr.ReadToEnd();
            }
            catch (ArgumentNullException)
            {
                origional.Body.Text += Environment.NewLine + Environment.NewLine + "Information Lost From The Combined Contacts Below:";
                foreach (string overflowedInfo in subOverflow)
                {
                    origional.Body.Text += Environment.NewLine + overflowedInfo;
                }
            }

            string firstHalf = "";
            string secondHalf = "";


            if (st.Contains("<html>"))
            {
                firstHalf = st.Split("</body>")[0];
                secondHalf = st.Split("</body>")[1];
                string turnIn = firstHalf;
                if (subOverflow.Count > 0)
                {
                    turnIn += Environment.NewLine + "<div>Information Lost From The Combined Contacts Below: " + origional.DisplayName + " and " + second.DisplayName + "</div>";
                    foreach (string overflowedInfo in subOverflow)
                    {
                        turnIn += Environment.NewLine + "<div>" + overflowedInfo + "</div>";
                    }
                    origional.Body.Text = turnIn + Environment.NewLine + "<div>End Of Lost Info</div>" + Environment.NewLine + "</body>" + Environment.NewLine + secondHalf;
                }




            }
            else
            {
                origional.Body.Text += Environment.NewLine + Environment.NewLine + "Information Lost From The Combined Contacts Below:";
                foreach (string overflowedInfo in subOverflow)
                {
                    origional.Body.Text += Environment.NewLine + overflowedInfo;
                }
                origional.Body.Text += Environment.NewLine + "End Of Lost Info";
            }*/
            return origional;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            Folder myFolder;
            List<Contact> items = new List<Contact>();
            ExchangeService _service = new ExchangeService
            {
                Credentials = new WebCredentials(textBoxEmail.Text, textBoxPassword.Text)
            };
            _service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            PropertySet propset = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Body);
            propset.RequestedBodyType = BodyType.Text;
            Folder msgF = Folder.Bind(_service, WellKnownFolderName.MsgFolderRoot);
            FolderView folderView = new FolderView(int.MaxValue);
            folderView.OffsetBasePoint = OffsetBasePoint.Beginning;
            folderView.PropertySet = new PropertySet(FolderSchema.DisplayName, FolderSchema.Id);
            FindFoldersResults folderResults = _service.FindFolders(msgF.Id, folderView);
            foreach (Folder folder in folderResults)
            {
                if (folder.DisplayName == textBox1.Text)
                {
                    myFolder = folder;
                    ItemView itemView = new ItemView(int.MaxValue, 0);
                    foreach (Contact item in myFolder.FindItems(itemView))
                    {
                        items.Add(item);
                        item.Load(propset);
                    }
                    break;
                }
            }


            EmailAddressKey[] emailKeys = { EmailAddressKey.EmailAddress1, EmailAddressKey.EmailAddress2, EmailAddressKey.EmailAddress3 };

            Dictionary<string, List<Contact>> contacts = new Dictionary<string, List<Contact>>();

            /*List<string> emails = new List<string>();

            ContactGroup contact = new ContactGroup(new Contact(_service), emails);

            List<Contact> toDel = new List<Contact>();*/

            //Make dictionary contain 1 email and a list of contacts. Every email gets it's own dictionary entry, but if ANY email is shared, the contact list is combined (they become the same list, changing one would change the other)
            List<List<Contact>> contactLists = new List<List<Contact>>();
            bool hasContactList;
            int contactListId = 0;
            string? shared1 = null;
            string? shared2 = null;
            foreach (Contact c in items)
            {
                shared1 = null;
                shared2 = null;
                hasContactList = false;
                foreach (EmailAddressKey key in emailKeys)
                {
                    if (c.EmailAddresses.Contains(key))
                    {
                        if (c.EmailAddresses[key].Address != null)
                        {
                            if (contacts.ContainsKey(c.EmailAddresses[key].Address))
                            {
                                if (contacts[c.EmailAddresses[key].Address].Contains(c) == false)
                                {
                                    if (hasContactList == true)
                                    {
                                        foreach (Contact con in contactLists[contactListId])
                                        {
                                            if (contacts[c.EmailAddresses[key].Address].Contains(con) == false)
                                            {
                                                contacts[c.EmailAddresses[key].Address].Add(con);
                                            }
                                        }
                                        foreach (KeyValuePair<string, List<Contact>> entry in contacts)
                                        {
                                            if (entry.Value == contactLists[contactListId])
                                            {
                                                contacts[entry.Key] = contacts[c.EmailAddresses[key].Address];
                                            }
                                        }
                                        if (shared1 != null)
                                        {
                                            contacts[shared1] = contacts[c.EmailAddresses[key].Address];
                                            if (shared2 != null)
                                            {
                                                contacts[shared2] = contacts[c.EmailAddresses[key].Address];
                                            }
                                        }

                                        contactLists.RemoveAt(contactListId);
                                        contactListId = contactLists.IndexOf(contacts[c.EmailAddresses[key].Address]);
                                    }
                                    else
                                    {
                                        contacts[c.EmailAddresses[key].Address].Add(c);
                                        contactListId = contactLists.IndexOf(contacts[c.EmailAddresses[key].Address]);
                                        hasContactList = true;
                                        if (shared1 == null)
                                        {
                                            shared1 = c.EmailAddresses[key].Address;
                                        }
                                        else
                                        {
                                            shared2 = c.EmailAddresses[key].Address;
                                        }
                                    }
                                }
                            }
                            else if (hasContactList == true)
                            {
                                contacts.Add(c.EmailAddresses[key].Address, contactLists[contactListId]);
                            }
                            else
                            {
                                contacts.Add(c.EmailAddresses[key].Address, new List<Contact> { c });
                                contactLists.Add(contacts[c.EmailAddresses[key].Address]);
                                contactListId = contactLists.Count - 1;
                                hasContactList = true;
                            }
                        }
                    }
                }
            }

            foreach (var dictEntry in contacts)
            {
                for (int i = 1; i < dictEntry.Value.Count;)
                {
                    combineContacts(dictEntry.Value[0], dictEntry.Value[i]);
                    dictEntry.Value[i].Delete(DeleteMode.SoftDelete);
                    dictEntry.Value.RemoveAt(i);
                }
                if (dictEntry.Value.Count != 0)
                {
                    dictEntry.Value[0].Update(ConflictResolutionMode.AutoResolve);
                }
                dictEntry.Value.Clear();
            }


            /*foreach (Contact c in items)
            {
                foreach (EmailAddressKey key in emailKeys)
                {
                    if (c.EmailAddresses.Contains(key))
                    {
                        if (c.EmailAddresses[key].Address != null)
                        {
                            contact.Addresses.Add(c.EmailAddresses[key].Address.ToString());
                        }
                    }
                }


                if (contact.Addresses.Count > 0)
                {
                    contact.Addresses.Sort();
                    if (contacts.ContainsKey(contact.Addresses[0]) == false)
                    {
                        Contact x = c;
                        List<string> y = new List<string>(contact.Addresses);
                        contacts.Add(contact.Addresses[0], new ContactGroup(x, y));
                    }
                    else
                    {
                        foreach (string add in contact.Addresses)
                        {
                            if (contacts[contact.Addresses[0]].Addresses.Contains(add) == false)
                            {
                                contacts[contact.Addresses[0]].Addresses.Add(add);
                            }
                        }
                        contacts[contact.Addresses[0]].Con = combineContacts(contacts[contact.Addresses[0]].Con, c);
                        toDel.Add(c);
                    }
                }
                contact.Addresses.Clear();
            }


            //Check for matching name or no? CURRENTLY DOES NOT
            foreach (KeyValuePair<string, ContactGroup> c in contacts)
            {
                if (c.Value.Combined == false)
                {
                    foreach (KeyValuePair<string, ContactGroup> c2 in contacts)
                    {
                        if (c2.Value.Combined == false && c.Key != c2.Key)
                        {
                            for (int i = 0; i < c.Value.Addresses.Count; i++)
                            {
                                if (c2.Value.Addresses.Contains(c.Value.Addresses[i]))
                                {
                                    c.Value.Con = combineContacts(c.Value.Con, c2.Value.Con);
                                    c2.Value.Combined = true;
                                    foreach (string add in c2.Value.Addresses)
                                    {
                                        if (c.Value.Addresses.Contains(add) == false)
                                        {
                                            c.Value.Addresses.Add(add);
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            foreach (KeyValuePair<string, ContactGroup> c in contacts)
            {
                c.Value.Con.Update(ConflictResolutionMode.AutoResolve);
            }

            for (int i = toDel.Count - 1; i >= 0; i--)
            {
                toDel[i].Delete(DeleteMode.SoftDelete);
            }*/
        }
    }
}