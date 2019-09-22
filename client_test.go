package grouppolicy

import (
	"eps/yzlog"
	"testing"
)

func TestRemoveGPO(t *testing.T) {
	cli := Client{}
	cli.RemoveGPO("gponame", "", "", false)
	cli.RemoveGPO("", "111111", "shinian.com", true)
	cli.RemoveGPO("aaa", "bbb", "dododod", true)
}

func TestGetAllGPO(t *testing.T) {

	cli := Client{}
	gpos, err := cli.GetAllGPO("shinian.com")
	if err != nil {
		t.Fatal(err)
	}

	for _, gpo := range gpos {
		yzlog.Infof("%+v", gpo)
	}
}

func TestGetGPO(t *testing.T) {
	cli := Client{}
	gpo, err := cli.GetGPO("testGPO", "1111", "shinian.com")
	if err != nil {
		t.Fatal(err)
	}
	yzlog.Infof("%+v", gpo)
}

func TestNewGPO(t *testing.T) {

	cli := NewClient("192.168.15.154")
	gpo, err := cli.NewGPO("testgpo", "aaaa", "shinian.com")
	if err != nil {
		t.Fatal(err)
	}
	yzlog.Infof("%+v", gpo)
}
