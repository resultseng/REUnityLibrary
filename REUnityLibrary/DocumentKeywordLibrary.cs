using System;
using System.Collections.Generic;
using Hyland.Unity;

namespace REUnityLibrary
{
    public class DocumentKeywordLibrary
    {
        private Hyland.Unity.Application _app;
        private Hyland.Unity.Document _doc;

        private List<ReadOnlyKeyword> _docKeywords;

        /// <summary>
        /// DocKeywordsLibrary. This library contains several methods to manipulate the keywords on a given document. 
        /// </summary>
        public DocumentKeywordLibrary(Hyland.Unity.Application app, Hyland.Unity.Document doc)
        {
            _app = app;
            _doc = doc;

            //no need to set diagnostic level, it comes in with app parameter

            //call this to initialize _docKeywords
            GetAllKeywordsList();
        }

        #region GetSingleKeywordValueByName
        /// <summary>
        /// Get Keyword value by Name.  If multiple keywords exist on the document, only the first value is returned by this function.
        /// This and GetKeywordValueFromList produce the same resulting keyword value, but GetSingleKeywordValueByName can be used to determine if the keyword exists.
        /// If you do not care if the keyword exists and both functions return empty string, use GetKeywordValueFromList because it is faster.
        /// </summary>
        /// <param name="keywordName"></param>
        /// <param name="keywordValue"></param>
        /// <returns>Returns true if the keyword was found, false if not</returns>
        public bool GetSingleKeywordValueByName(string keywordName, out string keywordValue)
        {
            bool retVal = false;
            try
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("enter GetSingleKeywordValueByName Keyword: {0}", keywordName));

                // Retrieve the keyword type for the keyword to add
                KeywordType keywordType = _app.Core.KeywordTypes.Find(keywordName);

                // Ensure keyword type was found
                if (keywordType == null)
                {
                    throw new Exception(string.Format("Keyword Type '{0}' not found", keywordName));
                }

                KeywordRecord keywordRecord = _doc.KeywordRecords.Find(keywordType);

                // Ensure keyword record was found
                if (keywordRecord == null)
                {
                    throw new Exception(string.Format(
                        "Keyword Record not found with keyword type: '{0}' on document: {1})",
                        keywordName, _doc.ID));
                }

                // Look for the first instance of a keyword in the keyword record
                //      with the specified keyword type
                Keyword keyword = keywordRecord.Keywords.Find(keywordType);

                // Ensure keyword was found
                if (keyword == null)
                {
                    throw new Exception(string.Format(
                        "Keyword not found with keyword type: '{0}' on keyword record: {1})",
                        keywordName, keywordRecord.ID));
                }

                if (keyword.IsBlank)
                    keywordValue = "";
                else
                    keywordValue = keyword.Value.ToString();

                retVal = true;

                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose,
                        string.Format("GetSingleKeywordValueByName(DocID: {0} Keyword: {1}) = {2}", _doc.ID, keywordName, keywordValue));
            }
            catch (UnityAPIException uae)
            {
                keywordValue = "";
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, "GetSingleKeywordValueByName unity error: " + uae.Message);
            }
            catch (Exception ex)
            {
                keywordValue = "";
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, "GetSingleKeywordValueByName error: " + ex.Message);
            }
            return retVal;
        }
        #endregion //GetSingleKeywordValueByName

        #region DocumentTypeContainsKeywordType
        /// <summary>
        /// DocumentTypeContainsKeywordType
        /// </summary>
        /// <param name="keywordTypeName"></param>
        /// <returns>true or false</returns>
        public bool DocumentTypeContainsKeywordType(string keywordTypeName)
        {
            bool result = false;

            string output = string.Format("Enter DocumentTypeContainsKeywordType: keywordTypeName={0}", keywordTypeName);
            _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, output);
            output = string.Format("DocTypeID={0}, DocTypeID={1}", _doc.DocumentType.ID, _doc.DocumentType.Name);
            _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, output);

            try
            {
                KeywordType keywordType = _app.Core.KeywordTypes.Find(keywordTypeName);

                // Ensure keyword type was found
                if (keywordType == null)
                {
                    _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, "keyword type does not exist in system");
                    return false;
                }

                //KeywordRecord keywordRecord = doc.KeywordRecords.Find(keywordType);
                KeywordRecord keywordRecord = _doc.KeywordRecords.Find(keywordType);

                // Ensure keyword record was found
                if (keywordRecord == null)
                {
                    _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, "keyword type exists in system, but not on this doc type");
                    return false;
                }

                result = true;

            }
            catch (UnityAPIException uae)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, "DocumentContainsKeywordType unity error: " + uae.Message);
            }
            catch (Exception ex)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, "DocumentContainsKeywordType error: " + ex.Message);
            }

            return result;

        }
        #endregion //DocumentTypeContainsKeywordType

        #region GetKeywordValuesListByName
        /// <summary>
        /// Get all keyword values for a single keyword type by name.
        /// </summary>
        /// <param name="keywordTypeName"></param>
        /// <returns></returns>
        public List<string> GetKeywordValuesListByName(string keywordTypeName)
        {
            //this function will return all values for a keywordtype in a list<string>
            _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, "Enter GetKeywordValuesListByName(" + keywordTypeName + ")");

            List<string> retVal = new List<string>();

            try
            {
                KeywordType keyType = _app.Core.KeywordTypes.Find(keywordTypeName);
                if (keyType == null)
                {
                    throw new Exception(String.Format("Keyword Type '{0}' not found in OnBase", keywordTypeName));
                }

                KeywordRecordList keyRecords = _doc.KeywordRecords.FindAll(keyType);
                if (keyRecords == null)
                {
                    throw new Exception(String.Format("Keyword Type '{0}' not found on document", keywordTypeName));
                }

                KeywordList keyList;
                foreach (KeywordRecord keyRecord in keyRecords)
                {
                    keyList = keyRecord.Keywords.FindAll(keyType);

                    foreach (Keyword kw in keyList)
                    {
                        if (!kw.IsBlank)
                            retVal.Add(kw.ToString());
                        else
                            retVal.Add("");
                    }
                }

            }
            catch (UnityAPIException uae)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, "GetKeywordValuesListByName unity error: " + uae.Message + Environment.NewLine + "Keyword type requested: " + keywordTypeName);
            }
            catch (Exception ex)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, "GetKeywordValuesListByName error: " + ex.Message + Environment.NewLine + "Keyword type requested: " + keywordTypeName);
            }
            return retVal;
        }
        #endregion //GetKeywordValuesListByName

        #region GetAllKeywordsList
        /// <summary>
        /// Retrieve a list object of ReadOnlyKeywords for all keywords on document
        /// </summary>
        private void GetAllKeywordsList()
        {
            _docKeywords = new List<ReadOnlyKeyword>();

            ReadOnlyKeyword detailRecord;

            try
            {
                foreach (KeywordRecord keyRecord in _doc.KeywordRecords)
                {
                    foreach (Keyword keyword in keyRecord.Keywords)
                    {
                        detailRecord = new ReadOnlyKeyword();
                        detailRecord.RecordType = keyRecord.KeywordRecordType.RecordType;
                        detailRecord.Name = keyword.KeywordType.Name;

                        if (!keyword.IsBlank)
                        {
                            detailRecord.Value = keyword.ToString();
                        }
                        else
                        {
                            detailRecord.Value = "";
                        }

                        _docKeywords.Add(detailRecord);
                    }
                }
            }
            catch (UnityAPIException uae)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, "GetAllKeywordsList unity error: " + uae.Message);
            }
            catch (Exception ex)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, "GetAllKeywordsList error: " + ex.Message);
            }

        }
        #endregion //GetAllKeywordsList

        /// <summary>
        /// A list of all keywords on the document presented as ReadOnlyKeyword objects
        /// </summary>
        public List<ReadOnlyKeyword> AllKeywordsList { get { return _docKeywords; } }

        #region GetKeywordValueFromList
        /// <summary>
        /// Get a single keyword value from the list of all keywords on the document
        /// </summary>
        /// <returns>Returns empty string if the keyword was not found</returns>
        public string GetKeywordValueFromList(string Name)
        {
            string value = "";
            try
            {
                if (_docKeywords != null)
                {
                    ReadOnlyKeyword kw = _docKeywords.Find(x => x.Name == Name);
                    value = kw.Value;
                }
            }
            catch (UnityAPIException uae)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, "GetKeywordValueFromList unity error: " + uae.Message + Environment.NewLine + "Keyword type requested: " + Name);
            }
            catch (Exception ex)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, "GetKeywordValueFromList error: " + ex.Message + Environment.NewLine + "Keyword type requested: " + Name);
            }

            _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, "GetKeywordValueFromList( " + Name + " ) = " + value);
            return value;

        }
        #endregion //GetKeywordValueFromList

        #region EditOrAddKeyword
        /// <summary>
        /// EditOrAddKeyword to a document
        /// </summary>
        /// <param name="keywordName"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public bool EditOrAddKeyword(string keywordName, string newValue)
        {
            _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("enter EditOrAddKeyword (DocID: {0}, Keyword: {1}, Value: {2})", _doc.ID, keywordName, newValue));

            bool retVal = false;

            try
            {
                KeywordType searchKeywordType = _app.Core.KeywordTypes.Find(keywordName);
                // Ensure keyword type was found
                if (searchKeywordType == null)
                {
                    throw new Exception(string.Format("Keyword Type '{0}' not found", keywordName));
                }

                Keyword newKeyword;
                switch (searchKeywordType.DataType)
                {
                    case KeywordDataType.AlphaNumeric:
                        newKeyword = searchKeywordType.CreateKeyword(newValue);
                        break;
                    case KeywordDataType.Currency:
                        newKeyword = searchKeywordType.CreateKeyword(Convert.ToDecimal(newValue));
                        break;
                    case KeywordDataType.Date:
                    case KeywordDataType.DateTime:
                        newKeyword = searchKeywordType.CreateKeyword(Convert.ToDateTime(newValue));
                        break;
                    case KeywordDataType.FloatingPoint:
                        newKeyword = searchKeywordType.CreateKeyword(Convert.ToDouble(newValue));
                        break;
                    case KeywordDataType.Numeric20:
                    case KeywordDataType.Numeric9:
                        newKeyword = searchKeywordType.CreateKeyword(Convert.ToInt64(newValue));
                        break;
                    default:
                        newKeyword = null;
                        break;
                }

                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, "  val=" + newKeyword.Value.ToString());

                // Lock document:  using statement handles disposal of DocumentLock object and releases lock
                //  Alternatively, DocumentLock.Release may be called to release the lock
                using (DocumentLock documentLock = _doc.LockDocument())
                {
                    // Ensure lock was obtained
                    if (documentLock.Status != DocumentLockStatus.LockObtained)
                    {
                        throw new Exception("Document lock not obtained");
                    }

                    KeywordModifier keyModifier = _doc.CreateKeywordModifier();

                    // Retrieve all keyword records on the document
                    //      which contains the keyword type
                    // Note: the Find method also has overrides that allow developers 
                    //      to search based on Keyword, or KeywordRecordType
                    KeywordRecordList keywordRecordList = _doc.KeywordRecords.FindAll(searchKeywordType);

                    _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, "  Keyword record count=" + keywordRecordList.Count.ToString());

                    if (keywordRecordList.Count == 0)
                    {
                        //add keyword
                        keyModifier.AddKeyword(newKeyword);
                        _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("  Keyword added: Type={0}, Value={1}", keywordName, newValue));
                    }
                    else
                    {
                        // Iterate through keyword record list 
                        foreach (KeywordRecord keywordRecord in keywordRecordList)
                        {
                            // Retrieve a list of keywords in the keyword record with the specified keyword type
                            KeywordList keywordList = keywordRecord.Keywords.FindAll(searchKeywordType);
                            // Iterate through keyword items in keyword list
                            foreach (Keyword keyword in keywordList)
                            {
                                if (keyword.IsBlank)
                                {
                                    keyModifier.AddKeyword(newKeyword);
                                    _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("  Keyword was blank, updated: Type={0}, New Value={1}", keywordName, newValue));
                                }
                                else
                                {
                                    keyModifier.UpdateKeyword(keyword, newKeyword);
                                    _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("  Keyword updated: Type={0}, Old Value={1}, New Value={2}", keywordName, keyword.Value.ToString(), newValue));
                                }
                            }
                        }
                    }
                    keyModifier.ApplyChanges();
                }
                retVal = true;
            }
            catch (UnityAPIException uae)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, uae);
            }
            catch (Exception ex)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, ex);
            }
            return retVal;
        }
        #endregion //EditOrAddKeyword

        #region AddKeyword
        /// <summary>
        /// Add a Keyword to a document; this will add a new instance of a keyword if one already exists, unlike EditOrAddKeyword which will perform an update if an instance already exists.
        /// </summary>
        /// <param name="keywordName"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public bool AddKeyword(string keywordName, string newValue)
        {
            _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("enter AddKeyword (DocID: {0}, Keyword: {1}, Value: {2})", _doc.ID, keywordName, newValue));

            bool retVal = false;

            try
            {
                KeywordType searchKeywordType = _app.Core.KeywordTypes.Find(keywordName);
                // Ensure keyword type was found
                if (searchKeywordType == null)
                {
                    throw new Exception(string.Format("Keyword Type '{0}' not found", keywordName));
                }

                Keyword newKeyword;
                switch (searchKeywordType.DataType)
                {
                    case KeywordDataType.AlphaNumeric:
                        newKeyword = searchKeywordType.CreateKeyword(newValue);
                        break;
                    case KeywordDataType.Currency:
                        newKeyword = searchKeywordType.CreateKeyword(Convert.ToDecimal(newValue));
                        break;
                    case KeywordDataType.Date:
                    case KeywordDataType.DateTime:
                        newKeyword = searchKeywordType.CreateKeyword(Convert.ToDateTime(newValue));
                        break;
                    case KeywordDataType.FloatingPoint:
                        newKeyword = searchKeywordType.CreateKeyword(Convert.ToDouble(newValue));
                        break;
                    case KeywordDataType.Numeric20:
                    case KeywordDataType.Numeric9:
                        newKeyword = searchKeywordType.CreateKeyword(Convert.ToInt64(newValue));
                        break;
                    default:
                        newKeyword = null;
                        break;
                }

                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, "  val=" + newKeyword.Value.ToString());

                // Lock document:  using statement handles disposal of DocumentLock object and releases lock
                //  Alternatively, DocumentLock.Release may be called to release the lock
                using (DocumentLock documentLock = _doc.LockDocument())
                {
                    // Ensure lock was obtained
                    if (documentLock.Status != DocumentLockStatus.LockObtained)
                    {
                        throw new Exception("Document lock not obtained");
                    }

                    KeywordModifier keyModifier = _doc.CreateKeywordModifier();
                    keyModifier.AddKeyword(newKeyword);
                    keyModifier.ApplyChanges();
                }
                retVal = true;
            }
            catch (UnityAPIException uae)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, uae);
            }
            catch (Exception ex)
            {
                _app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, ex);
            }
            return retVal;
        }
        #endregion //AddKeyword
    }
}
