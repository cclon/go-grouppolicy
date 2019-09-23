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

type PSBoolean string

const (
	Unspecified PSBoolean = "Unspecified"
	No          PSBoolean = "No"
	Yes         PSBoolean = "Yes"
)

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

// RenameGPO
func (cli Client) RenameGPO(name, guid, targetName, domain string) error {
	if len(name) == 0 && len(guid) == 0 {
		return errors.New("you must point one of name or guid")
	}
	if len(targetName) == 0 {
		return errors.New("you must set a targetName")
	}

	cmd := `Rename-GPO`
	if len(guid) != 0 {
		cmd += fmt.Sprintf(` -Guid "%s"`, guid)
	} else {
		cmd += fmt.Sprintf(` "%s"`, name)
	}

	cmd += fmt.Sprintf(` -TargetName "%s"`, targetName)
	if len(domain) != 0 {
		cmd += fmt.Sprintf(` -Domain "%s"`, domain)
	}

	yzlog.Debug(cmd)
	_, _, err := runLocalPowershell(cmd)
	return err
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
	_, _, err := runLocalPowershell(cmd)
	return err
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
	return gpo, nil
}

// SetGPLink 设置GPO链接的属性(和New-GPLink参数一致)
func (cli Client) SetGPLink(name,
	target,
	domain string,
	enforced,
	linkEnabled PSBoolean) error {

	return cli.NewGPLink(name, target, domain, enforced, linkEnabled)
}

// NewGPLink 链接一个GPO到站点(site)，域名(Domain)或者组织单位(OU)
// name - 要进行链接的GPO名称
// target - 要连接到的目标对象，ou或者域等
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
	yzlog.Debug(cmd)

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

// BackupGPO 备份指定的
func (cli Client) BackupGPO(name, guid, path, comment string) error {

	if len(name) == 0 && len(guid) == 0 {
		return errors.New("you must point one of name or guid")
	}

	if len(path) == 0 {
		return errors.New("you must set Restore-GPO path")
	}

	cmd := `Backup-GPO`
	if len(comment) != 0 {
		cmd += fmt.Sprintf(`-Comment "%s"`, comment)
	}
	if len(guid) != 0 {
		cmd += fmt.Sprintf(` -Guid "%s"`, guid)
	} else {
		cmd += fmt.Sprintf(` "%s"`, name)
	}
	cmd += fmt.Sprintf(` -path "%s"`, path)

	yzlog.Debug(cmd)
	_, _, err := runLocalPowershell(cmd)
	return err
}

// ImportGPO 导入GPO配置，并覆盖已经存在GPO
// 注意： 目标GPO必须存在
func (cli Client) ImportGPO(backupGpoName, targetName, path string) error {

	if len(backupGpoName) == 0 || len(targetName) == 0 || len(path) == 0 {
		return errors.New("invalid paramer")
	}

	cmd := fmt.Sprintf(`Import-GPO -BackupGpoName "%s" -TargetName "%s" -path "%s"`,
		backupGpoName, targetName, path)
	yzlog.Debug(cmd)
	_, _, err := runLocalPowershell(cmd)
	return err
}

// RestoreGPO 还原指定备份GPO
func (cli Client) RestoreGPO(name, guid, path string) error {

	if len(name) == 0 && len(guid) == 0 {
		return errors.New("you must point one of name or guid")
	}

	if len(path) == 0 {
		return errors.New("you must set Restore-GPO path")
	}

	cmd := `Restore-GPO`
	if len(guid) != 0 {
		cmd += fmt.Sprintf(` -Guid "%s"`, guid)
	} else {
		cmd += fmt.Sprintf(` "%s"`, name)
	}
	cmd += fmt.Sprintf(` -path "%s"`, path)

	yzlog.Debug(cmd)
	_, _, err := runLocalPowershell(cmd)
	return err
}

// InvokeGpupdate 更新指定主机的组策略
func (cli Client) InvokeGpupdate(computer, target string) error {
	cmd := fmt.Sprintf(`Invoke-GPUpdate -Computer "%s" -Target "%s"`, computer, target)
	yzlog.Debug(cmd)
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
