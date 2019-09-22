package main

import (
	"eps/yzlog"
	"fmt"
	gp "go-grouppolicy"
)

func main() {

	fmt.Println(gp.IsGroupPolicyModuleInstalled())

	cli := gp.Client{}

	// 删除一个gplink
	err := cli.RemoveGPLink(`testgplink`, "", "ou=appinstall,dc=shinian,dc=com", "")
	if err != nil {
		yzlog.Error(err)
		return
	}

	// 删除一个gpo
	err = cli.RemoveGPO(`testgpo`, "", "", false)
	if err != nil {
		yzlog.Error(err)
		return
	}

	// 创建一个新的GPO
	gpo, err := cli.NewGPO(`testgpo`, "this is a test gpo", "")
	if err != nil {
		yzlog.Error(err)
		return
	}
	yzlog.Infof("%+v", gpo)

	// get gpo info
	gpo, err = cli.GetGPO(`testgpo`, "", "")
	if err != nil {
		yzlog.Error(err)
		return
	}
	yzlog.Infof("%+v", gpo)

	// create gplink
	err = cli.NewGPLink(`testgplink`, `ou=appinstall,dc=shinian,dc=com`, "", "", gp.Yes)
	if err != nil {
		yzlog.Error(err)
		return
	}

}
