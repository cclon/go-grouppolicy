package grouppolicy

import "fmt"

type GPOReportType string

const (
	GPOReportTypeXml  = GPOReportType("Xml")
	GPOReportTypeHtml = GPOReportType("Html")
)

type GPO struct {
	ID          string `json:"Id"`
	DisplayName string `json:"DisplayName"`
	Owner       string `json:"Owner"`
	DomainName  string `json:"DomainName"`
	// CreationTime     string
	// ModificationTime string
	// User             User
	// Computer         Computer
	// GpoStatus        int
	// WmiFilter        string
	Description string `json:"Description"`
}

func (gpo GPO) string() string {
	return ""
}

// GetReport
func (gpo GPO) GetReport(format GPOReportType) (string, error) {
	cmd := fmt.Sprintf(`Get-GPOReport -Guid %s -ReportType %s`, gpo.ID, format)
	stdout, _, err := runLocalPowershell(cmd)
	if err != nil {
		return "", err
	}
	return stdout, nil
}
