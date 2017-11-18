using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BordxGenerator.Model
{
    class ClaimBordx
    {
        public string InsuredId { get; set; }
        public string Insured { get; set; }
        public Address Address { get; set; }
        public Address FLAddress { get; set; }
        public string ClaimantId { get; set; }
        public string Claimant { get; set; }
        public string PolicyNumber { get; set; }
        public DateTime EffectiveDate { get; set; }
        public DateTime ExpirationDate { get; set; }
        public Int64 Year { get; set; }
        public DateTime LossDateFrom { get; set; }
        public DateTime LossDateTo { get; set; }
        public string ClaimNumber { get; set; }
        public string FileReportedToUnderwritter { get; set; }
        public bool DenialFile { get; set; }
        public DateTime DateClaimDenial { get; set; }
        public string DenialReason { get; set; }
        public bool Complaint { get; set; }
        public bool Litigation { get; set; }
        public string LossDescription { get; set; }
        public string LossLocation { get; set; }
        public string OriginalCurrency { get; set; }
        public string SettlementCurrency { get { return "USD"; } }
        public double AmountClaimed { get; set; }
        public double AmountPaid { get; set; }
        public double RecoveryThisMonth { get; set; }
        public double FeesPaid { get; set; }
        public double IndemnityReserve { get; set; }
        public double AdjusterReserve { get; set; }
        public DateTime DateClaimMade { get; set; }
        public DateTime DateClaimNotified { get; set; }
        public DateTime DateClaimOpened { get; set; }
        public DateTime DateCoverageAgreed { get; set; }
        public DateTime DateClaimDenied { get; set; }
        public string ReasonForDenial { get; set; }
        public DateTime DateClaimAmountAgreed { get; set; }
        public DateTime DateClaimPaid { get; set; }
        public DateTime DateFeesPaid { get; set; }
        public bool Status { get; set; }
        public DateTime DateFileClosed { get; set; }
        public DateTime DateSubrogationStarted { get; set; }
        public DateTime DateFileReopened { get; set; }
        public DateTime DateClaimWithdrawn { get; set; }
        public string Adjuster { get; set; }
        public string AdjusterBranch { get; set; }

    }
}
