using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseDocumentReferenceClassLibrary
{
    public class DocumentReferenceItem
    {
        public string DocPrefix;
        public string DocRefType;
        public string DocReference;
        public string DocRefOther;
        public string DocumentType;
        public string DocumentName1;
        public string DocumentNo;
        public string DocumentRef;
        public string DocVerNo;
        public string District;
        public string ElecRef;
        public string ElecType;
        public string VerStatus;
        public string VerType;

        public DocumentReferenceItem()
        {

        }

        public DocumentReferenceItem(DocumentReferenceService.DocumentReferenceDTO docRefDto)
        {
            District = docRefDto.dstrctCode;

            DocPrefix = docRefDto.docPrefix;
            DocRefType = docRefDto.docRefType;
            DocReference = docRefDto.docReference;
            DocRefOther = docRefDto.docRefOther;

            DocumentType = docRefDto.documentType;
            DocumentName1 = docRefDto.documentName1;
            DocumentNo = docRefDto.documentNo;
            DocumentRef = docRefDto.documentRef;

            DocVerNo = docRefDto.docVerNo;
            ElecRef = docRefDto.elecRef;
            ElecType = docRefDto.elecType;
            VerStatus = docRefDto.verStatus;
            VerType = docRefDto.verType;

        }
    }
}
