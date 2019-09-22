package grouppolicy

import (
	"encoding/json"
	"eps/yzlog"
	"errors"
	"fmt"
	"regexp"
	"strings"

	ps "github.com/ao-com/go-powershell"
	"github.com/ao-com/go-powershell/backend"
)

type Client struct {
}

// IsGroupPolicyModuleInstalled
// 判断是否安装了powershell的组策略模块
func IsGroupPolicyModuleInstalled() (bool, error) {

	cmd := "if (Get-Module -List grouppolicy) {'true'}"
	stdout, _, err := runLocalPowershell(cmd)
	if err != nil {
		return false, err
	}
	return strings.Contains(stdout, "true"), nil
}

type GPError struct {
	ErrorId      string // 错误类型
	CategoryInfo string // 错误明细
}

func (e GPError) Error() string {
	return fmt.Sprintf("ErrorId[%s] CategoryInfo[%s]",
		e.ErrorId,
		e.CategoryInfo)
}

var regexpCategoryInfo = regexp.MustCompile(`CategoryInfo[\s]*: (.+)\r`)
var regexpErrorId = regexp.MustCompile(`FullyQualifiedErrorId : (.+)\r`)

// runLocalPowershell
func runLocalPowershell(cmd string) (string, string, error) {

	back := &backend.Local{}
	shell, err := ps.New(back)
	if err != nil {
		return "", "", err
	}
	defer shell.Exit()

	stdout, stderr, err := shell.Execute(cmd)

	var gpErr GPError

	// 通过正则提出错误信息
	if err != nil {
		if msg := regexpCategoryInfo.FindStringSubmatch(stderr); len(msg) == 2 {
			gpErr.CategoryInfo = msg[1]
		}

		if msg := regexpErrorId.FindStringSubmatch(stderr); len(msg) == 2 {
			gpErr.ErrorId = msg[1]
		}

		if idx := strings.Index(gpErr.CategoryInfo, gpErr.ErrorId); idx != -1 {
			gpErr.CategoryInfo = gpErr.CategoryInfo[0 : idx-2]
		}

		err = gpErr
	}

	return stdout, stderr, err
}

// NewGPO 新建一个GPO
func (cli Client) NewGPO(name, comment, domain string) (*GPO, error) {

	if len(name) == 0 {
		return nil, errors.New("name could not null")
	}

	cmd := fmt.Sprintf(`New-GPO "%s"`, name)
	if len(comment) != 0 {
		cmd += fmt.Sprintf(` -Comment "%s"`, comment)
	}
	if len(domain) != 0 {
		cmd += fmt.Sprintf(` -Domain "%s"`, domain)
	}

	cmd += ` | ConvertTo-Json`

	yzlog.Debug(cmd)

	stdout, _, err := runLocalPowershell(cmd)
	if err != nil {
		return nil, err
	}

	gpo := &GPO{}
	err = json.Unmarshal([]byte(stdout), gpo)
	if err != nil {
		return nil, err
	}
	return gpo, err
}

// RemoveGPO 删除一个GPO
// 可以通过name或者GUID进行删除，如果都为空，返回错误，如果都指定，则使用guid，忽略name
func (cli Client) RemoveGPO(name, guid, domain string, keeplinks bool) error {

	cmd := `remove-gpo`
	if len(name) == 0 && len(guid) == 0 {
		return errors.New("you must point one of name or guid")
	}

	if len(guid) != 0 {
		cmd += fmt.Sprintf(` -guid "%s"`, guid)
	} else if len(name) != 0 {
		cmd += fmt.Sprintf(` "%s"`, name)
	}

	if len(domain) != 0 {
		cmd += fmt.Sprintf(` -domain "%s"`, domain)
	}

	if keeplinks {
		cmd += ` -keeplinks`
	}

	yzlog.Debug(cmd)
	//_, _, err := runLocalPowershell(cmd)
	return nil
}

// GetAllGPO 返回所有GPO列表
func (cli Client) GetAllGPO(domain string) (gpos []*GPO, err error) {

	cmd := `Get-GPO -All`
	if len(domain) != 0 {
		cmd += fmt.Sprintf(` -Domain "%s"`, domain)
	}

	cmd += ` | ConvertTo-Json`

	yzlog.Debug(cmd)

	stdout, stderr, err := runLocalPowershell(cmd)
	_ = stderr
	if err != nil {
		return nil, err
	}

	gpos = make([]*GPO, 0)
	err = json.Unmarshal([]byte(stdout), &gpos)
	if err != nil {
		return nil, err
	}
	return gpos, nil
}

// 返回指定参数条件的GPO列表
// 可以通过name或者GUID进行查询GPO信息，如果都为空，返回错误，如果都指定，则使用guid，忽略name
// you can see how to use from http://go.microsoft.com/fwlink/?LinkId=216700
func (cli Client) GetGPO(name, guid, domain string) (*GPO, error) {

	if len(name) == 0 && len(guid) == 0 {
		return nil, errors.New("you must point one of name or guid")
	}
	cmd := `Get-GPO`
	if len(guid) != 0 {
		cmd += fmt.Sprintf(` "%s"`, guid)
	} else {
		cmd += fmt.Sprintf(` "%s"`, name)
	}

	if len(domain) != 0 {
		cmd += fmt.Sprintf(` -Domain "%s"`, domain)
	}

	cmd += ` | ConvertTo-Json`
	stdout, _, err := runLocalPowershell(cmd)
	if err != nil {
		return nil, err
	}

	gpo := &GPO{}
	err = json.Unmarshal([]byte(stdout), gpo)
	if err != nil {
		return nil, err
	}
	return gpo, nil
}

// RestoreGPO 还原指定备份GPO
func (cli Client) RestoreGPO(name string, path string) error {

	cmd := fmt.Sprintf(`Restore-GPO -Name "%s" -Path "%s"`, name, path)
	_, _, err := runLocalPowershell(cmd)
	return err
}

// SetGPLink 设置GPO链接的属性
func (cli Client) SetGPLink() error {
	return errors.New("unknown")
}

type PSBoolean string

const (
	Unspecified PSBoolean = "Unspecified"
	No          PSBoolean = "No"
	Yes         PSBoolean = "Yes"
)

// NewGPLink 链接一个GPO到站点(site)，域名(Domain)或者组织单位(OU)
func (cli Client) NewGPLink(name,
	target,
	domain string,
	enforced, linkEnabled PSBoolean) error {

	if len(name) == 0 {
		return errors.New("you must point a name")
	}

	if len(target) == 0 {
		return errors.New("you must set GPLink target")
	}

	cmd := fmt.Sprintf(`New-GPLink "%s"  -Target "%s"`, name, target)

	if len(domain) != 0 {
		cmd += fmt.Sprintf(` -Domain "%s"`, domain)
	}
	if len(enforced) != 0 {
		cmd += fmt.Sprintf(` -Enforced "%s"`, enforced)
	}
	if len(linkEnabled) != 0 {
		cmd += fmt.Sprintf(` -LinkEnabled "%s"`, linkEnabled)
	}

	_, _, err := runLocalPowershell(cmd)
	return err
}

// RemoveGPLink 删除一个GPO链接
func (cli Client) RemoveGPLink(name, guid, target, domain string) error {

	if len(name) == 0 && len(guid) == 0 {
		return errors.New("you must point one of name or guid")
	}
	if len(target) == 0 {
		return errors.New("you must set Remove-GPLink target")
	}

	cmd := `Remove-GPLink`
	if len(guid) != 0 {
		cmd += fmt.Sprintf(` -Guid "%s"`, guid)
	} else {
		cmd += fmt.Sprintf(` "%s"`, name)
	}

	if len(domain) != 0 {
		cmd += fmt.Sprintf(` -Domain "%s"`, domain)
	}
	cmd += fmt.Sprintf(` -Target "%s"`, target)

	yzlog.Debug(cmd)

	_, _, err := runLocalPowershell(cmd)
	return err
}

// InvokeGpupdate 更新指定主机的组策略
func (cli Client) InvokeGpupdate(computer, target string) error {
	cmd := fmt.Sprintf(`Invoke-GPUpdate -Computer "%s" -Target "%s"`, computer, target)
	_, _, err := runLocalPowershell(cmd)
	return err
}

// RemoveGPRegistryValue
func (cli Client) RemoveGPRegistryValue() {

}

// GetGPRegistryValue
func (cli Client) GetGPRegistryValue() {
}

// SetGPRegistryValue
func (cli Client) SetGPRegistryValue() {
}
