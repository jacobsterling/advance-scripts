# -*- coding: utf-8 -*-
"""
Created on Wed Jan 19 16:13:57 2022

@author: jacob.sterling
"""

import os
from datetime import date, datetime

from zcrmsdk.src.com.zoho.crm.api import HeaderMap, ParameterMap
from zcrmsdk.src.com.zoho.crm.api.attachments import Attachment
from zcrmsdk.src.com.zoho.crm.api.layouts import Layout
from zcrmsdk.src.com.zoho.crm.api.record import *
from zcrmsdk.src.com.zoho.crm.api.record import Record as ZCRMRecord
from zcrmsdk.src.com.zoho.crm.api.tags import Tag
from zcrmsdk.src.com.zoho.crm.api.users import User
from zcrmsdk.src.com.zoho.crm.api.util import Choice, StreamWrapper

class Record(object):
    @staticmethod
    def get_record('Contacts', 1171216000172203413, r'C:\Users\jacob.sterling\OneDrive - advance.online\Operations\PythonSDK\jfjkfg.txt'):

        """
        This method is used to get a single record of a module with ID and print the response.
        :param module_api_name: The API Name of the record's module.
        :param record_id: The ID of the record to be obtained.
        :param destination_folder: The absolute path of the destination folder to store the downloaded attachment Record
        """

        """
        example
        module_api_name = 'Leads'
        record_id = 3477061000006603276
        """

        # Get instance of RecordOperations Class
        record_operations = RecordOperations()

        # Get instance of ParameterMap Class
        param_instance = ParameterMap()

        # Possible parameters for Get Record operation
        param_instance.add(GetRecordParam.cvid, record_id)

        param_instance.add(GetRecordParam.approved, 'true')

        param_instance.add(GetRecordParam.converted, 'both')

        fields = ['id', 'company']

        for field in fields:
            param_instance.add(GetRecordParam.fields, field)

        start_date_time = datetime(2020, 1, 1, 10, 10, 10)

        param_instance.add(GetRecordParam.startdatetime, start_date_time)

        end_date_time = datetime(2020, 7, 7, 10, 10, 10)

        param_instance.add(GetRecordParam.enddatetime, end_date_time)

        param_instance.add(GetRecordParam.territory_id, record_id)

        param_instance.add(GetRecordParam.include_child, 'true')

        param_instance.add(GetRecordParam.uid, record_id)

        # Get instance of HeaderMap Class
        header_instance = HeaderMap()

        # Possible headers for Get Record operation
        header_instance.add(GetRecordHeader.if_modified_since, datetime.now())

        # header_instance.add(GetRecordHeader.x_external, "Leads.External")

        # Call getRecord method that takes param_instance, header_instance, module_api_name and record_id as parameter
        response = record_operations.get_record(record_id, module_api_name, param_instance, header_instance)

        if response is not None:

            # Get the status code from response
            print('Status Code: ' + str(response.get_status_code()))

            if response.get_status_code() in [204, 304]:
                print('No Content' if response.get_status_code() == 204 else 'Not Modified')
                return

            # Get object from response
            response_object = response.get_object()

            if response_object is not None:

                # Check if expected ResponseWrapper instance is received.
                if isinstance(response_object, ResponseWrapper):

                    # Get the list of obtained Record instances
                    record_list = response_object.get_data()

                    for record in record_list:
                        # Get the ID of each Record
                        print("Record ID: " + str(record.get_id()))

                        # Get the createdBy User instance of each Record
                        created_by = record.get_created_by()

                        # Check if created_by is not None
                        if created_by is not None:
                            # Get the Name of the created_by User
                            print("Record Created By - Name: " + created_by.get_name())

                            # Get the ID of the created_by User
                            print("Record Created By - ID: " + str(created_by.get_id()))

                            # Get the Email of the created_by User
                            print("Record Created By - Email: " + created_by.get_email())

                        # Get the CreatedTime of each Record
                        print("Record CreatedTime: " + str(record.get_created_time()))

                        if record.get_modified_time() is not None:
                            # Get the ModifiedTime of each Record
                            print("Record ModifiedTime: " + str(record.get_modified_time()))

                        # Get the modified_by User instance of each Record
                        modified_by = record.get_modified_by()

                        # Check if modified_by is not None
                        if modified_by is not None:
                            # Get the Name of the modified_by User
                            print("Record Modified By - Name: " + modified_by.get_name())

                            # Get the ID of the modified_by User
                            print("Record Modified By - ID: " + str(modified_by.get_id()))

                            # Get the Email of the modified_by User
                            print("Record Modified By - Email: " + modified_by.get_email())

                        # Get the list of obtained Tag instance of each Record
                        tags = record.get_tag()

                        if tags is not None:
                            for tag in tags:
                                # Get the Name of each Tag
                                print("Record Tag Name: " + tag.get_name())

                                # Get the Id of each Tag
                                print("Record Tag ID: " + str(tag.get_id()))

                        # To get particular field value
                        print("Record Field Value: " + str(record.get_key_value('Last_Name')))

                        print('Record KeyValues: ')

                        key_values = record.get_key_values()

                        for key_name, value in key_values.items():

                            if isinstance(value, list):

                                if len(value) > 0:

                                    if isinstance(value[0], FileDetails):
                                        file_details = value

                                        for file_detail in file_details:
                                            # Get the Extn of each FileDetails
                                            print("Record FileDetails Extn: " + file_detail.get_extn())

                                            # Get the IsPreviewAvailable of each FileDetails
                                            print("Record FileDetails IsPreviewAvailable: " + str(file_detail.get_is_preview_available()))

                                            # Get the DownloadUrl of each FileDetails
                                            print("Record FileDetails DownloadUrl: " + file_detail.get_download_url())

                                            # Get the DeleteUrl of each FileDetails
                                            print("Record FileDetails DeleteUrl: " + file_detail.get_delete_url())

                                            # Get the EntityId of each FileDetails
                                            print("Record FileDetails EntityId: " + file_detail.get_entity_id())

                                            # Get the Mode of each FileDetails
                                            print("Record FileDetails Mode: " + file_detail.get_mode())

                                            # Get the OriginalSizeByte of each FileDetails
                                            print("Record FileDetails OriginalSizeByte: " + file_detail.get_original_size_byte())

                                            # Get the PreviewUrl of each FileDetails
                                            print("Record FileDetails PreviewUrl: " + file_detail.get_preview_url())

                                            # Get the FileName of each FileDetails
                                            print("Record FileDetails FileName: " + file_detail.get_file_name())

                                            # Get the FileId of each FileDetails
                                            print("Record FileDetails FileId: " + file_detail.get_file_id())

                                            # Get the AttachmentId of each FileDetails
                                            print("Record FileDetails AttachmentId: " + file_detail.get_attachment_id())

                                            # Get the FileSize of each FileDetails
                                            print("Record FileDetails FileSize: " + file_detail.get_file_size())

                                            # Get the CreatorId of each FileDetails
                                            print("Record FileDetails CreatorId: " + file_detail.get_creator_id())

                                            # Get the LinkDocs of each FileDetails
                                            print("Record FileDetails LinkDocs: " + file_detail.get_link_docs())

                                    elif isinstance(value[0], Reminder):
                                        reminders = value

                                        for reminder in reminders:
                                            # Get the Reminder Period
                                            print("Reminder Period: " + reminder.get_period())

                                            # Get the Reminder Unit
                                            print("Reminder Unit: " + reminder.get_unit())

                                    elif isinstance(value[0], Choice):
                                        choice_list = value

                                        print(key_name)

                                        print('Values')

                                        for choice in choice_list:
                                            print(choice.get_value())

                                    elif isinstance(value[0], Participants):
                                        participants = value

                                        for participant in participants:
                                            print("Record Participants Name: ")

                                            print(participant.get_name())

                                            print("Record Participants Invited: " + str(participant.get_invited()))

                                            print("Record Participants Type: " + participant.get_type())

                                            print("Record Participants Participant: " + participant.get_participant())

                                            print("Record Participants Status: " + participant.get_status())

                                    elif isinstance(value[0], InventoryLineItems):
                                        product_details = value

                                        for product_detail in product_details:
                                            line_item_product = product_detail.get_product()

                                            if line_item_product is not None:
                                                print("Record ProductDetails LineItemProduct ProductCode: " + line_item_product.get_product_code())

                                                print("Record ProductDetails LineItemProduct Currency: " + line_item_product.get_currency())

                                                print("Record ProductDetails LineItemProduct Name: " + line_item_product.get_name())

                                                print("Record ProductDetails LineItemProduct Id: " + line_item_product.get_id())

                                            print("Record ProductDetails Quantity: " + str(product_detail.get_quantity()))

                                            print("Record ProductDetails Discount: " + product_detail.get_discount())

                                            print("Record ProductDetails TotalAfterDiscount: " + str(product_detail.get_total_after_discount()))

                                            print("Record ProductDetails NetTotal: " + str(product_detail.get_net_total()))

                                            if product_detail.get_book() is not None:
                                                print("Record ProductDetails Book: " + str(product_detail.get_book()))

                                            print("Record ProductDetails Tax: " + str(product_detail.get_tax()))

                                            print("Record ProductDetails ListPrice: " + str(product_detail.get_list_price()))

                                            print("Record ProductDetails UnitPrice: " + str(product_detail.get_unit_price()))

                                            print("Record ProductDetails QuantityInStock: " + str(product_detail.get_quantity_in_stock()))

                                            print("Record ProductDetails Total: " + str(product_detail.get_total()))

                                            print("Record ProductDetails ID: " + product_detail.get_id())

                                            print("Record ProductDetails ProductDescription: " + product_detail.get_product_description())

                                            line_taxes = product_detail.get_line_tax()

                                            for line_tax in line_taxes:
                                                print("Record ProductDetails LineTax Percentage: " + str(line_tax.get_percentage()))

                                                print("Record ProductDetails LineTax Name: " + line_tax.get_name())

                                                print("Record ProductDetails LineTax Id: " + line_tax.get_id())

                                                print("Record ProductDetails LineTax Value: " + str(line_tax.get_value()))

                                    elif isinstance(value[0], Tag):
                                        tags = value

                                        if tags is not None:
                                            for tag in tags:
                                                print("Record Tag Name: " + tag.get_name())

                                                print("Record Tag ID: ")

                                                print(tag.get_id())

                                    elif isinstance(value[0], PricingDetails):
                                        pricing_details = value

                                        for pricing_detail in pricing_details:
                                            print("Record PricingDetails ToRange: " + str(pricing_detail.get_to_range()))

                                            print("Record PricingDetails Discount: " + str(pricing_detail.get_discount()))

                                            print("Record PricingDetails ID: " + pricing_detail.get_id())

                                            print("Record PricingDetails FromRange: " + str(pricing_detail.get_from_range()))

                                    elif isinstance(value[0], ZCRMRecord):
                                        record_list = value

                                        for each_record in record_list:
                                            for key, val in each_record.get_key_values().items():
                                                print(str(key) + " : " + str(val))

                                    elif isinstance(value[0], LineTax):
                                        line_taxes = value

                                        for line_tax in line_taxes:
                                            print("Record LineTax Percentage: " + str(
                                                line_tax.get_percentage()))

                                            print("Record LineTax Name: " + line_tax.get_name())

                                            print("Record LineTax Id: " + line_tax.get_id())

                                            print("Record LineTax Value: " + str(line_tax.get_value()))

                                    elif isinstance(value[0], Comment):
                                        comments = value

                                        for comment in comments:
                                            print("Comment-ID: " + comment.get_id())

                                            print("Comment-Content: " + comment.get_comment_content())

                                            print("Comment-Commented_By: " + comment.get_commented_by())

                                            print("Comment-Commented Time: " + str(comment.get_commented_time()))

                                    elif isinstance(value[0], Attachment):
                                        attachments = value

                                        for attachment in attachments:
                                            # Get the ID of each attachment
                                            print('Record Attachment ID : ' + str(attachment.get_id()))

                                            # Get the owner User instance of each attachment
                                            owner = attachment.get_owner()

                                            # Check if owner is not None
                                            if owner is not None:
                                                # Get the Name of the Owner
                                                print("Record Attachment Owner - Name: " + owner.get_name())

                                                # Get the ID of the Owner
                                                print("Record Attachment Owner - ID: " + owner.get_id())

                                                # Get the Email of the Owner
                                                print("Record Attachment Owner - Email: " + owner.get_email())

                                            # Get the modified time of each attachment
                                            print("Record Attachment Modified Time: " + str(attachment.get_modified_time()))

                                            # Get the name of the File
                                            print("Record Attachment File Name: " + attachment.get_file_name())

                                            # Get the created time of each attachment
                                            print("Record Attachment Created Time: " + str(attachment.get_created_time()))

                                            # Get the Attachment file size
                                            print("Record Attachment File Size: " + str(attachment.get_size()))

                                            # Get the parentId Record instance of each attachment
                                            parent_id = attachment.get_parent_id()

                                            if parent_id is not None:
                                                # Get the parent record Name of each attachment
                                                print(
                                                    "Record Attachment parent record Name: " + parent_id.get_key_value("name"))

                                                # Get the parent record ID of each attachment
                                                print("Record Attachment parent record ID: " + parent_id.get_id())

                                            # Check if the attachment is Editable
                                            print("Record Attachment is Editable: " + str(attachment.get_editable()))

                                            # Get the file ID of each attachment
                                            print("Record Attachment File ID: " + str(attachment.get_file_id()))

                                            # Get the type of each attachment
                                            print("Record Attachment File Type: " + str(attachment.get_type()))

                                            # Get the seModule of each attachment
                                            print("Record Attachment seModule: " + str(attachment.get_se_module()))

                                            # Get the modifiedBy User instance of each attachment
                                            modified_by = attachment.get_modified_by()

                                            # Check if modifiedBy is not None
                                            if modified_by is not None:
                                                # Get the Name of the modifiedBy User
                                                print("Record Attachment Modified By - Name: " + modified_by.get_name())

                                                # Get the ID of the modifiedBy User
                                                print("Record Attachment Modified By - ID: " + modified_by.get_id())

                                                # Get the Email of the modifiedBy User
                                                print("Record Attachment Modified By - Email: " + modified_by.get_email())

                                            # Get the state of each attachment
                                            print("Record Attachment State: " + attachment.get_state())

                                            # Get the modifiedBy User instance of each attachment
                                            created_by = attachment.get_created_by()

                                            # Check if created_by is not None
                                            if created_by is not None:
                                                # Get the Name of the modifiedBy User
                                                print("Record Attachment Created By - Name: " + created_by.get_name())

                                                # Get the ID of the modifiedBy User
                                                print("Record Attachment Created By - ID: " + created_by.get_id())

                                                # Get the Email of the modifiedBy User
                                                print("Record Attachment Created By - Email: " + created_by.get_email())

                                            # Get the linkUrl of each attachment
                                            print("Record Attachment LinkUrl: " + str(attachment.get_link_url()))

                                    else:
                                        print(key_name)

                                        for each_value in value:
                                            print(str(each_value))

                            elif isinstance(value, User):
                                print("Record " + key_name + " User-ID: " + str(value.get_id()))

                                print("Record " + key_name + " User-Name: " + value.get_name())

                                print("Record " + key_name + " User-Email: " + value.get_email())

                            elif isinstance(value, Layout):
                                print(key_name + " ID: " + str(value.get_id()))

                                print(key_name + " Name: " + value.get_name())

                            elif isinstance(value, ZCRMRecord):
                                print(key_name + " Record ID: " + str(value.get_id()))

                                print(key_name + " Record Name: " + value.get_key_value('name'))

                            elif isinstance(value, Choice):
                                print(key_name + " : " + value.get_value())

                            elif isinstance(value, RemindAt):
                                print(key_name + " : " + value.get_alarm())

                            elif isinstance(value, RecurringActivity):
                                print(key_name)

                                print("RRULE: " + value.get_rrule())

                            elif isinstance(value, Consent):
                                print("Record Consent ID: " + str(value.get_id()))

                                # Get the createdBy User instance of each Record
                                created_by = value.get_created_by()

                                # Check if created_by is not None
                                if created_by is not None:
                                    # Get the Name of the created_by User
                                    print("Record Consent Created By - Name: " + created_by.get_name())

                                    # Get the ID of the created_by User
                                    print("Record Consent Created By - ID: " + created_by.get_id())

                                    # Get the Email of the created_by User
                                    print("Record Consent Created By - Email: " + created_by.get_email())

                                # Get the CreatedTime of each Record
                                print("Record Consent CreatedTime: " + str(value.get_created_time()))

                                if value.get_modified_time() is not None:
                                    # Get the ModifiedTime of each Record
                                    print("Record Consent ModifiedTime: " + str(value.get_modified_time()))

                                # Get the Owner User instance of the Consent
                                owner = value.get_owner()

                                if owner is not None:
                                    # Get the Name of the Owner User
                                    print("Record Consent Created By - Name: " + owner.get_name())

                                    # Get the ID of the Owner User
                                    print("Record Consent Created By - ID: " + owner.get_id())

                                    # Get the Email of the Owner User
                                    print("Record Consent Created By - Email: " + owner.get_email())

                                print("Record Consent ContactThroughEmail: " + str(value.get_contact_through_email()))

                                print("Record Consent ContactThroughSocial: " + str(value.get_contact_through_social()))

                                print("Record Consent ContactThroughSurvey: " + str(value.get_contact_through_survey()))

                                print("Record Consent ContactThroughPhone: " + str(value.get_contact_through_phone()))

                                print("Record Consent MailSentTime: " + str(value.get_mail_sent_time()))

                                print("Record Consent ConsentDate: " + str(value.get_consent_date()))

                                print("Record Consent ConsentRemarks: " + value.get_consent_remarks())

                                print("Record Consent ConsentThrough: " + value.get_consent_through())

                                print("Record Consent DataProcessingBasis: " + value.get_data_processing_basis())

                                # To get custom values
                                print("Record Consent Lawful Reason: " + str(value.get_key_value("Lawful_Reason")))

                            elif isinstance(value, dict):
                                for key, val in value.items():
                                    print(key + " : " + str(val))

                            else:
                                print(key_name + " : " + str(value))

                # Check if expected FileBodyWrapper instance is received.
                elif isinstance(response_object, FileBodyWrapper):

                    # Get StreamWrapper instance from the returned FileBodyWrapper instance
                    stream_wrapper = response_object.get_file()

                    # Construct the file name by joining the destinationFolder and the name from StreamWrapper instance
                    file_name = os.path.join(destination_folder, stream_wrapper.get_name())

                    # Open the destination file where the file needs to be written in 'wb' mode
                    with open(file_name, 'wb') as f:
                        # Get the stream from StreamWrapper instance
                        for chunk in stream_wrapper.get_stream():
                            f.write(chunk)

                        f.close()

                # Check if the request returned an exception
                elif isinstance(response_object, APIException):
                    # Get the Status
                    print("Status: " + response_object.get_status().get_value())

                    # Get the Code
                    print("Code: " + response_object.get_code().get_value())

                    print("Details")

                    # Get the details dict
                    details = response_object.get_details()

                    for key, value in details.items():
                        print(key + ' : ' + str(value))

                    # Get the Message
                    print("Message: " + response_object.get_message().get_value())
                    
#get_record('Contacts', 1171216000172203413, r'C:\Users\jacob.sterling\OneDrive - advance.online\Operations\PythonSDK\jfjkfg.txt')