using System;
using Radar.Core.Models.Batch;

namespace ImportPOC2.Utils
{
    public class Validation
    {
        public static void AddValidationError(Batch batch, string criteriaCode, string info, long productId, string xid)
        {
            //TODO: criteria code will not always correctly map to field codes
            //TODO: where did "ILUV" error code come from? 

            batch.BatchErrorLogs.Add(new BatchErrorLog
            {
                FieldCode = criteriaCode,
                ErrorMessageCode = "ILUV",
                AdditionalInfo = info,
                ProductId = productId,
                ExternalProductId = xid
            });
        }
    }
}
