package grouppolicy

import (
	"encoding/json"
	"errors"
	"fmt"
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

// runLocalPowershell
func runLocalPowershell(cmd string) (stdout string, stderr string, err error) {

	back := &backend.Local{}
	shell, err := ps.New(back)
	if err != nil {
		return "", "", err
	}
	defer shell.Exit()

	fmt.Printf("%s\n", cmd)

	return shell.Execute(cmd)
}

// NewGPO 新建一个GPO
func (cli Client) NewGPO(name string, options map[string]interface{}) (*GPO, error) {

	var optionStr string
	if options != nil {
		for k, v := range options {
			switch strings.ToLower(k) {
			case "comment": // GPO描述
				optionStr += fmt.Sprintf(` -Comment "%s"`, v)
			case "domain": // 域名
				optionStr += fmt.Sprintf(` -Domain "%s"`, v)
			case "server": // 远端主机
				optionStr += fmt.Sprintf(` -Server "%s"`, v)
			case "startergponame":
				optionStr += fmt.Sprintf(` -StarterName "%s"`, v)
			case "startergpoguid":
				optionStr += fmt.Sprintf(` -StarterGpoGuid %s`, v)
			}
		}
	}

	cmd := fmt.Sprintf(`New-GPO "%s"`, name)
	if len(optionStr) > 0 {
		cmd += optionStr
	}
	cmd += `|  ConvertTo-Json`

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
func (cli Client) RemoveGPO(options map[string]interface{}) error {

	var optionStr string
	if options != nil {

		for k, v := range options {
			switch strings.ToLower(k) {
			case "server":
				optionStr += fmt.Sprintf(` -Server "%s"`, v)
			case "guid":
				optionStr += fmt.Sprintf(` -Guid "%s"`, v)
			case "name":
				optionStr += fmt.Sprintf(` -Name "%s"`, v)
			case "domain":
				optionStr += fmt.Sprintf(` -Domain "%s"`, v)
			case "keeplinks":
				optionStr += " -KeepLinks"
			}
		}
	}

	cmd := fmt.Sprintf(`Remove-GPO %s`, optionStr)
	_, _, err := runLocalPowershell(cmd)
	return err
}

//
func (cli Client) GetAllGPO(optionals map[string]interface{}) (gpos []*GPO, err error) {

	var optionalStr string
	cmd := `Get-GPO -All`

	if optionals != nil {
		for k, v := range optionals {
			switch strings.ToLower(k) {
			case "server":
				optionalStr += fmt.Sprintf(` -Server "%s"`, v)
			case "domain":
				optionalStr += fmt.Sprintf(` -Domain "%s"`, v)
			}
		}
	}

	if len(optionalStr) > 0 {
		cmd += optionalStr
	}

	cmd += ` | ConvertTo-Json`
	stdout, _, err := runLocalPowershell(cmd)
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
// options may have some options:
// Guid/Name/Domain/Server,All
// you can see how to use from http://go.microsoft.com/fwlink/?LinkId=216700
func (cli Client) GetGPO(nameOrGuid string, optionals map[string]interface{}) (*GPO, error) {

	var optionalStr string
	if optionals != nil {
		for k, v := range optionals {
			switch strings.ToLower(k) {
			case "server":
				optionalStr += fmt.Sprintf(` -Server "%s"`, v)
			case "domain":
				optionalStr += fmt.Sprintf(` -Domain "%s"`, v)
			}
		}
	}

	cmd := fmt.Sprintf(`Get-GPO %s`, nameOrGuid)
	if len(optionalStr) > 0 {
		cmd += optionalStr
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

// NewGPLink 链接一个GPO到站点(site)，域名(Domain)或者组织单位(OU)
func (cli Client) NewGPLink(srcGpoOption map[string]interface{}, target string) error {

	srcGpoOptionStr := ""
	if srcGpoOption != nil {
		for k, v := range srcGpoOption {
			switch strings.ToLower(k) {
			case "server":
				srcGpoOptionStr += fmt.Sprintf(` -Server "%s"`, v)
			case "guid":
				srcGpoOptionStr += fmt.Sprintf(` -Name "%s"`, v)
			case "name":
				srcGpoOptionStr += fmt.Sprintf(` -Guid "%s"`, v)
			case "domain":
				srcGpoOptionStr += fmt.Sprintf(` -Domain "%s"`, v)
			}
		}
	}

	cmd := fmt.Sprintf(`New-GPLink %s -Target "%s"`, srcGpoOptionStr, target)
	_, _, err := runLocalPowershell(cmd)
	return err
}

// RemoveGPLink 删除一个GPO链接
func (cli Client) RemoveGPLink(srcGpoOption map[string]interface{}, target string) error {

	optionStr := ""
	if srcGpoOption != nil {
		for k, v := range srcGpoOption {
			switch strings.ToLower(k) {
			case "server":
				optionStr += fmt.Sprintf(` -Server "%s"`, v)
			case "guid":
				optionStr += fmt.Sprintf(` -Name "%s"`, v)
			case "name":
				optionStr += fmt.Sprintf(` -Guid "%s"`, v)
			case "domain":
				optionStr += fmt.Sprintf(` -Domain "%s"`, v)
			}
		}
	}

	cmd := fmt.Sprintf(`Remove-GPLink %s -Target "%s"`, optionStr, target)
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
