type AccountMap = {
    [accountName: string]: {
        exportId: string;
        importId: string;
        markup: number;
    };
};

const AccountMapping: AccountMap = {
    "Alice and Ames": {
        exportId: "849524242",
        importId: "950371911",
        markup: 1.25
    },
    "Asher Golf": {
        exportId: "848827078",
        importId: "966528006",
        markup: 1.2
    },
    "Can-Am Consulting": {
        exportId: "849605448",
        importId: "963163949",
        markup: 1.25
    },
    "Cocojojo": {
        exportId: "849589139",
        importId: "962173127",
        markup: 1.25
    },
    "Dixxon Flannel": {
        exportId: "849088566",
        importId: "967230500",
        markup: 1.25
    },
    "Enlinx": {
        exportId: "849185490",
        importId: "968243585",
        markup: 1.5
    },
    "GG Design": {
        exportId: "849605422",
        importId: "963163936",
        markup: 1.25
    },
    "Gigi Pip": {
        exportId: "849227792",
        importId: "960167968",
        markup: 1.25
    },
    "Knife Aid": {
        exportId: "849583409",
        importId: "962146963",
        markup: 1.25
    },
    "Lizzy James": {
        exportId: "849036031",
        importId: "964210657",
        markup: 1.25
    },
    "Nominal": {
        exportId: "849165450",
        importId: "968174160",
        markup: 1.25
    },
    "Refyne": {
        exportId: "849605381",
        importId: "963163626",
        markup: 1.25
    },
    "Socks for Animals": {
        exportId: "849205420",
        importId: "969570499",
        markup: 1.1
    },
    "Transmark Sales": {
        exportId: "849146765",
        importId: "969595968",
        markup: 1.25
    },
    "Village Hat Shop": {
        exportId: "849136625",
        importId: "969570600",
        markup: 1.2
    },
    "Woodies": {
        exportId: "849056620",
        importId: "965281193",
        markup: 1.1
    }
};

export default AccountMapping;
