        /// <summary>
        /// Sends emails to <paramref name="emailAddresses"/> in parallel
        /// </summary>
        /// <param name="emailAddresses">The recipients' email addresses</param>
        /// <param name="allAttachmentURLs">URLs to the file attachments</param>
        public void Orchestrate(string[] emailAddresses, string[] allAttachmentURLs)
        {
            string sub = subjectLine.ToString();
            string body = File.ReadAllText(bodyLink.ToString());

            //Use -1 for MaxDegreeOfParallelism to set a NO limit to the degree of concurrency
            var options = new ParallelOptions { MaxDegreeOfParallelism = -1 };


            //sharedLocals is a threadsafe variable where each parallel iteration saves its own data
            ConcurrentBag<SharedLocal> sharedLocals = new ConcurrentBag<SharedLocal>();

            //start parallel ForEachLoop on the emailAddresses
                Parallel.ForEach(emailAddresses, options, email =>
                {
                    //each parallel iteration has its own sharedLocal variable
                    SharedLocal sharedLocal = new SharedLocal();

                    //this iteration will save its first data to its own sharedLocal
                    sharedLocal.Counter = 1;

                    string readline = email;
                    string[] parts = readline.Split('|');
                    string Name = parts[2].ToString();
                    string toAddress = parts[3].ToString().Trim().Replace(" ", "");
                    string attachment = parts[0].ToString().Trim();

                    if (toAddress != "")
                    {

                        //find the attachment file associated with this email address
                        long index = Array.FindIndex<string>(allAttachmentURLs, x => x.Contains(attachment));

                        if (index != -1)
                        {

                            //update a property in this iteration's own sharedLocal variable
                            sharedLocal.SBReport = allAttachmentURLs[index];


                            attachment = allAttachmentURLs[index];


                            //send mail
                            if (SendEmail(sub, body, toAddress, attachment, username, password, sharedLocal))
                            {
                                sharedLocal.SuccessCount++;
                                sharedLocal.SBSuccess = "Success!! >> " + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + ": " + readline;

                            }
                            else
                            {
                                sharedLocal.FailedCount++;
                                sharedLocal.SBFailed = "Failed!! >> " + DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + ": " + readline;
                            }

                        }

                        else
                        {

                            //report on the result in this iteration's own sharedLocal
                            sharedLocal.SBReportError = "Attachment Not Found For : " + readline;
                        }
                    }
                    else
                    {
                        //report on the result in this iteration's own sharedLocal
                        sharedLocal.SBReportError = "TO Address Issue For : " + readline;
                    }


                    //update the threadsafe ConcurrentBag with this iteration's own sharedLocal
                    sharedLocals.Add(sharedLocal);
                });
            

        }
