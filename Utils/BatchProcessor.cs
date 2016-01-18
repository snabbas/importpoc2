using System.Linq;
using ASI.Sugar.Collections;
using Radar.Core.Models.Batch;
using Radar.Data;

namespace ImportPOC2.Utils
{
    /// <summary>
    /// used to handle the logging of batch errors, batch retrieval, etc.
    /// </summary>
    public static class BatchProcessor
    {
        private static Batch _curBatch;
        public static UowPRODTask ProdTask { get; set; }
        public static string CurrentXid { get; set; }
        public static object logger { get; set; }

        /// <summary>
        /// Used when a lookup value is specified on sheet but not found in global lookup
        /// </summary>
        /// <param name="criteriaCode"></param>
        /// <param name="additionalInfo"></param>
        public static void AddLookupValidationError(string criteriaCode, string additionalInfo)
        {
            _curBatch.BatchErrorLogs.Add(
                new BatchErrorLog
                {
                    ExternalProductId = CurrentXid,
                    AdditionalInfo = additionalInfo,
                    ErrorMessageCode = "ILOK", //error text = "{0} value does not exist (for all lookups)"
                    FieldCode = string.Empty //TODO: fill this in via parameter somehow
                });

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="criteriaCode"></param>
        /// <param name="additionalInfo"></param>
        public static void AddIncorrectFormatError(string criteriaCode, string additionalInfo)
        {
            _curBatch.BatchErrorLogs.Add(
                new BatchErrorLog
                {
                    ExternalProductId = CurrentXid,
                    AdditionalInfo =  additionalInfo,
                    ErrorMessageCode = "IICF",// "{0} field is not in a correct format."
                    FieldCode = string.Empty //TODO: fill this in via parameter somehow
                });
        }

        public static void AddGenericFieldError(string fieldCode, string additionalInfo)
        {
            _curBatch.BatchErrorLogs.Add(
                new BatchErrorLog
                {
                    ExternalProductId = CurrentXid,
                    AdditionalInfo = additionalInfo,
                    ErrorMessageCode = "GENR",
                    FieldCode = fieldCode
                });
        }

        internal static bool SetCurrentBatch(long batchId)
        {
            var retVal = true;
            var batch = ProdTask.Batch.GetAllWithInclude(
                t => t.BatchDataSources,
                t => t.BatchErrorLogs,
                t => t.BatchProducts)
                .FirstOrDefault(b => b.BatchId == batchId);

            if (batch == null)
            {
                retVal = false;
            }
            else
            {
                _curBatch = batch;
            }
            return retVal;
        }

        public static int GetCompanyIdFromCurrentBatch()
        {
            int retVal = 0;
            if (_curBatch != null && _curBatch.CompanyId.HasValue)
            {
                retVal = (int)_curBatch.CompanyId.Value;
            }
            return retVal;
        }

        //ONLY FOR DEBUGGING
        //internal static void OutputBatchErrors()
        //{
        //    if (_curBatch != null && _curBatch.BatchErrorLogs.Any())
        //    {
        //        _curBatch.BatchErrorLogs.ForEach(e =>
        //        {
                    
        //        }
        //    );
        //    }
        //}

        internal static void OutputBatchErrors(log4net.ILog _log)
        {
            if (_curBatch != null && _curBatch.BatchErrorLogs.Any())
            {
                //_curBatch.BatchErrorLogs.ForEach(e => _log.DebugFormat("{0}:{1}", e.ErrorMessageCode, e.AdditionalInfo));
            }
        }
    }
}

/*
 *         private static void addValidationError(string criteriaCode, string info)
        {
            //TODO: criteria code will not always correctly map to field codes
            //TODO: where did "ILUV" error code come from? 

            _log.WarnFormat("Validation Error: {0}\r\n{1}", criteriaCode, info);

            _curBatch.BatchErrorLogs.Add(new BatchErrorLog
            {
                FieldCode = criteriaCode,
                ErrorMessageCode = "ILUV",
                AdditionalInfo = info,
                ProductId = _currentProduct.ID,
                ExternalProductId = _curXid
            });
        }

*/