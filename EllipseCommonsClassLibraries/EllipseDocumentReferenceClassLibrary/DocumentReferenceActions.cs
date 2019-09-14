using System;
using System.Linq;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseDocumentReferenceClassLibrary
{
    public static class DocumentReferenceActions
    {
        public static DocumentReferenceService.DocumentReferenceServiceResult CreateDocument(string urlService, DocumentReferenceService.OperationContext opContext, DocumentReferenceItem docRef)
        {
            try
            {
                if (!ValidateReferences(docRef))
                    throw new Exception("Referencia no válida. Verifique nuevamente los datos");
                //ejecuta las acciones del servicio
                var service = new DocumentReferenceService.DocumentReferenceService
                {
                    Url = urlService + "/DocumentReference"
                };

                //Instanciar el Contexto de Operación
                var request = new DocumentReferenceService.DocumentReferenceDTO();
                request.docPrefix = docRef.DocPrefix;
                request.docRefType = docRef.DocRefType;
                request.docReference = docRef.DocReference;
                request.docRefOther = docRef.DocRefOther;
                request.documentName1 = docRef.DocumentName1;
                request.documentNo = docRef.DocumentNo;
                request.documentRef = docRef.DocumentRef;
                request.documentType = docRef.DocumentType;
                request.docVerNo = docRef.DocVerNo;
                request.dstrctCode = opContext.district;
                request.elecRef = docRef.ElecRef;
                request.elecType = docRef.ElecType;
                request.verStatus = docRef.VerStatus;
                request.verType = docRef.VerType;

                var reply = service.createDoc(opContext, request);
                
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("DocumentReferenceActions:CreateDocument(String, DocumentReferenceService.OperationContext, DocumentReferenceItem)", ex.Message);
                throw;
            }
        }

        public static DocumentReferenceService.DocumentReferenceServiceResult DeleteDocument(string urlService, DocumentReferenceService.OperationContext opContext, DocumentReferenceItem docRef)
        {
            try
            {
                if (!ValidateReferences(docRef))
                    throw new Exception("Referencia no válida. Verifique nuevamente los datos");
                //ejecuta las acciones del servicio
                var service = new DocumentReferenceService.DocumentReferenceService
                {
                    Url = urlService + "/DocumentReference"
                };

                //Instanciar el Contexto de Operación

                var request = new DocumentReferenceService.DocumentReferenceDTO();
                request.docPrefix = docRef.DocPrefix;//
                request.docRefType = docRef.DocRefType;
                request.docReference = docRef.DocReference;
                request.docRefOther = docRef.DocRefOther;//
                request.documentName1 = docRef.DocumentName1;
                request.documentNo = docRef.DocumentNo;
                request.documentRef = docRef.DocumentRef;//
                request.documentType = docRef.DocumentType;
                request.docVerNo = docRef.DocVerNo;
                request.dstrctCode = opContext.district;
                request.elecRef = docRef.ElecRef;
                request.elecType = docRef.ElecType;
                request.verStatus = docRef.VerStatus;//
                request.verType = docRef.VerType;

                var reply = service.delete(opContext, request);
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("DocumentReferenceActions:DeleteDocument(String, DocumentReferenceService.OperationContext, DocumentReferenceItem)", ex.Message);
                throw;
            }
        }

        public static DocumentReferenceService.DocumentReferenceServiceResult LinkDocument(string urlService, DocumentReferenceService.OperationContext opContext, DocumentReferenceItem docRef)
        {
            try
            {
                if (!ValidateReferences(docRef))
                    throw new Exception("Referencia no válida. Verifique nuevamente los datos");
                //ejecuta las acciones del servicio
                var service = new DocumentReferenceService.DocumentReferenceService
                {
                    Url = urlService + "/DocumentReference"
                };
                //Instanciar el Contexto de Operación

                var request = new DocumentReferenceService.DocumentReferenceDTO();
                request.docPrefix = docRef.DocPrefix;//
                request.docRefType = docRef.DocRefType;
                request.docReference = docRef.DocReference;
                request.docRefOther = docRef.DocRefOther;//
                request.documentName1 = docRef.DocumentName1;
                request.documentNo = docRef.DocumentNo;
                request.documentRef = docRef.DocumentRef;//
                request.documentType = docRef.DocumentType;
                request.docVerNo = docRef.DocVerNo;
                request.dstrctCode = opContext.district;
                request.elecRef = docRef.ElecRef;
                request.elecType = docRef.ElecType;
                request.verStatus = docRef.VerStatus;//
                request.verType = docRef.VerType;

                var reply = service.linkDoc(opContext, request);
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("DocumentReferenceActions:DeleteDocument(String, DocumentReferenceService.OperationContext, DocumentReferenceItem)", ex.Message);
                throw;
            }
        }

        public static DocumentReferenceService.DocumentReferenceServiceResult UpdateDocument(string urlService, DocumentReferenceService.OperationContext opContext, DocumentReferenceItem docRef)
        {
            try
            {
                if (!ValidateReferences(docRef))
                    throw new Exception("Referencia no válida. Verifique nuevamente los datos");
                //ejecuta las acciones del servicio
                var service = new DocumentReferenceService.DocumentReferenceService
                {
                    Url = urlService + "/DocumentReference"
                };
                //Instanciar el Contexto de Operación

                var request = new DocumentReferenceService.DocumentReferenceDTO();
                request.docPrefix = ""+ docRef.DocPrefix;//
                request.docRefType = "" + docRef.DocRefType;
                request.docReference = "" + docRef.DocReference;
                request.docRefOther = "" + docRef.DocRefOther;//
                request.documentName1 = "" + docRef.DocumentName1;
                request.documentNo = "" + docRef.DocumentNo;
                request.documentRef = "" + docRef.DocumentRef;//
                request.documentType = "" + docRef.DocumentType;
                request.docVerNo = "" + docRef.DocVerNo;
                request.dstrctCode = "" + opContext.district;
                request.elecRef = "" + docRef.ElecRef;
                request.elecType = "" + docRef.ElecType;
                request.verStatus = "" + docRef.VerStatus;//
                request.verType = "" + docRef.VerType;
                var reply = service.update(opContext, request);
                var resp = new DocumentReferenceService.DocumentReferenceServiceResult();
                
                return reply;
            }
            catch (Exception ex)
            {
                Debugger.LogError("DocumentReferenceActions:DeleteDocument(String, DocumentReferenceService.OperationContext, DocumentReferenceItem)", ex.Message);
                throw;
            }
        }

        private static bool ValidateReferences(DocumentReferenceItem docRef)
        {
            //CN - Contact
            //CO - Contract
            //CU - Compatible Unit
            //CV - Contract Variation
            //DO - Document
            //DT - Defect
            //EG - EGI (Equipment Group Id)
            //EM - Employee Id
            //EQ - Equipment
            //ER - Emp Rehab
            //ET - Medical Tst
            //FP - Forward Purchase Agreement
            //GA - Global Pur Agreement
            //GI - Global Pur Agreement Item
            //GO - Supplier Offer
            //GQ - Request for Quote
            //IG - Isolation Guide
            //IN - Incident
            //IS - Isolation Sheet
            //IV - Invoice
            //JE - Job Estimate
            //LN - Location
            //PA - Part
            //PJ - Project
            //PO - Purchase Order
            //PR - Purchase Requisition
            //PS - Position
            //PV - Proj. Approval
            //RE - Issue Requisition
            //RO - Recommended Order
            //SC - Stock Code
            //SJ - Standard Job
            //SU - Supplier
            //TC - Training Course
            //TS - Training Session
            //VA - Contract Valuation
            //WC - Claim
            //WO - Work Order
            //WR - Work Request

            //TO DO
            if(docRef.Equals("WO"))
            {
                //TO DO
                
            }
            else if (docRef.Equals("WR"))
            {
                //TO DO
            }
            

            return true;
        }
        /// <summary>
        /// Crea un nuevo operador de contexto para los métodos de la clase
        /// </summary>
        /// <param name="district">string: Distrito donde se va a crear el contexto</param>
        /// <param name="position">string: Posición donde se va a crear el contexto</param>
        /// <param name="maxInstances">int: Número máximo de instancias</param>
        /// <param name="returnWarnings">bool: True no ignora las advertencias</param>
        /// <returns></returns>
        public static DocumentReferenceService.OperationContext GetDocRefOpContext(string district, string position, int maxInstances, bool returnWarnings)
        {
            var opContext = new DocumentReferenceService.OperationContext
            {
                district = district,
                position = position,
                maxInstances = maxInstances,
                maxInstancesSpecified = true,
                returnWarnings = returnWarnings,
                returnWarningsSpecified = true
            };

            return opContext;
        }

        public static DocumentReferenceService.OperationContext GetDocRefOpContext()
        {
            return new DocumentReferenceService.OperationContext();
        }
    }
}
