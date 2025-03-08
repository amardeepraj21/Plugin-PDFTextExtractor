using PDFTextExtractor;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IdentityModel.Metadata;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Crm.Sdk.Messages;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Linq;
using System.Reflection;

namespace PDFTextExtractor
{
    public class TextExtractor : PluginBase
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="PostOperationbdapp_parkingsyncprocessCreate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information.</param>
        public TextExtractor(string unsecure, string secure)
            : base(typeof(TextExtractor))
        {
            EmbeddedAssemblyResolver.AttachResolver();
            // TODO: Implement your custom configuration handling.
        }

        /// <summary>
        /// Main entry point for he business logic that the plug-in is to execute.
        /// </summary>
        /// <param name="localContext">The <see cref="localContext"/> which contains the
        /// <see cref="IPluginExecutionContext"/>,
        /// <see cref="IOrganizationService"/>
        /// and <see cref="ITracingService"/>
        /// </param>
        /// <remarks>
        /// </remarks>
        /// 
        protected override void ExecuteCdsPlugin(ILocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException(nameof(localContext));
            }

            IPluginExecutionContext context = localContext.PluginExecutionContext;
            IOrganizationService service = localContext.CurrentUserService;
            ITracingService tracingService = localContext.TracingService;

            //******************************//////////////////////////////////////////////////////////////////////////////////////////////////////////////
            string messageName = context.MessageName.ToLower();



            // Determine whether it is Create, Update, or Delete
            if (messageName == "create")
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity entity)
                {
                    if (entity.LogicalName != "ast_secondpdftextextractor")
                        throw new InvalidPluginExecutionException("Expected Logical Name not found i.e. ast_secondpdftextextractor");

                    if (entity.Contains("ast_name"))
                    {
                        // Get the base64 PDF bytes from the entity attribute
                        var pdfBase64 = entity["ast_name"] as string;

                        if (!string.IsNullOrEmpty(pdfBase64))
                        {
                            // Convert the base64 string to a byte array
                            //byte[] pdfBytes = Convert.FromBase64String(pdfBase64);

                            //Update Extracted Text
                            entity["ast_completepdfextractedtext"] = pdfBase64;

                            // Update the entity in Dataverse
                            service.Update(entity);
                        }
                        else
                        {
                            // Log error if the base64 is empty or null
                            throw new InvalidOperationException("PDF base64 bytes are null or empty.");
                        }
                    }
                }
            }
            else if (messageName == "update")
            {
                if (context.Depth > 1)
                {
                    return;
                }
                else
                {

                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity entity)
                    {
                        if (entity.LogicalName != "ast_secondpdftextextractor")
                            throw new InvalidPluginExecutionException("Expected Logical Name not found i.e. ast_secondpdftextextractor");

                        if (entity.Contains("ast_pdfuploaded"))
                        {
                            string fileColumnName = "ast_secpdfattachment";  // Attachment Column name ast_secpdfattachment
                            Guid recordId = entity.Id;
                            string entityName = entity.LogicalName;

                            // Step 1: Initialize file download
                            var initRequest = new InitializeFileBlocksDownloadRequest
                            {
                                FileAttributeName = fileColumnName,
                                Target = new EntityReference(entityName, recordId)
                            };

                            var initResponse = (InitializeFileBlocksDownloadResponse)service.Execute(initRequest);

                            // Step 2: Download the file block
                            var downloadRequest = new DownloadBlockRequest
                            {
                                FileContinuationToken = initResponse.FileContinuationToken
                            };

                            var downloadResponse = (DownloadBlockResponse)service.Execute(downloadRequest);

                            // Convert bytes to Base64
                            string base64String = Convert.ToBase64String(downloadResponse.Data);


                            if (!string.IsNullOrEmpty(base64String))
                            {
                                // Convert the base64 string to a byte array
                                byte[] pdfBytes = Convert.FromBase64String(base64String);

                                // Extract text from the second last page of the PDF and fetch email addresses
                                string extractedText = ExtractTextFromSecondLastPage(pdfBytes);

                                //Update Extracted Text
                                entity["ast_completepdfextractedtext"] = extractedText;


                                // Fetch email addresses from the extracted text
                                string emailAddresses = ExtractEmailAddresses(extractedText);

                                // Assuming 'ast_emailaddress' is where the email addresses should go
                                entity["ast_extractedemailaddress"] = emailAddresses;

                                // Update the entity in Dataverse
                                service.Update(entity);
                            }
                            else
                            {
                                // Log error if the base64 is empty or null
                                throw new InvalidOperationException("PDF base64 bytes are null or empty.");
                            }
                        }
                    }
                }
            }

        }

        //Private methods

        private string ExtractTextFromSecondLastPage(byte[] pdfBytes)
        {
            try
            {
                using (var memoryStream = new MemoryStream(pdfBytes))
                using (var reader = new PdfReader(memoryStream))
                using (var pdfDoc = new PdfDocument(reader))
                {
                    int totalPages = pdfDoc.GetNumberOfPages();

                    // Check if there's at least two pages in the document
                    if (totalPages < 2)
                    {
                        return string.Empty; // Return empty if the second-to-last page doesn't exist
                    }

                    int secondLastPageIndex = totalPages - 1; // Second last page

                    var page = pdfDoc.GetPage(secondLastPageIndex);
                    var strategy = new LocationTextExtractionStrategy();
                    return PdfTextExtractor.GetTextFromPage(page, strategy);
                }
            }
            catch (Exception ex)
            {
                // Handle and log any errors during text extraction
                throw new InvalidOperationException("Error extracting text from the second last page: " + ex.Message, ex);
            }
        }

        private string ExtractEmailAddresses(string extractedText)
        {
            try
            {
                // Check if the extracted text is null or empty
                if (string.IsNullOrEmpty(extractedText))
                {
                    return null; // Return null if no text is extracted
                }

                // Define a simple regex pattern for matching email addresses
                string emailPattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}";

                // Find all matches for the email pattern
                var emailMatches = Regex.Matches(extractedText, emailPattern);

                // If no email addresses are found, return null
                if (emailMatches.Count == 0)
                {
                    return null;
                }

                // Concatenate all email addresses with a semicolon
                var emailList = emailMatches.Cast<Match>()
                                           .Select(m => m.Value)
                                           .ToList();

                return string.Join(";", emailList);
            }
            catch (Exception ex)
            {
                // Handle and log any errors during email extraction
                throw new InvalidOperationException("Error extracting email addresses: " + ex.Message, ex);
            }
        }


    }
}


