##### 3.7.3.1.1 PDUs are described as octets or groups of octets   PDU描述为八位字节或八位字节组
a) Groups of octets separated by a comma appear in the order they are transferred If optional octets are not present the following octets appear without a gap

`以逗号分隔的八位字节组按传输顺序显示。如果不存在可选的八位位组，则以下八位位元将显示为无间隙`

b) If octets or groups of octets are grouped within “{ }” the order is arbitrary

`如果八位字节或八位字节组在“｛｝”内分组，则顺序是任意的`

c) If octets or groups of octets are marked with “*” they may appear more than once. 
If it is used within a “{ }” section they may appear mixed with other octets or group of octets of this section.

`如果八位字节或八位字节组标有“*”，它们可能会出现多次。`

`如果在“｛｝”部分中使用，它们可能会与该部分的其他八位字节或一组八位字节混合。`

d) Octets can be grouped or values can be assigned within “( )”

`八位字节可以分组，也可以在“（）”内赋值`

e) If octets or groups of octets are grouped within “[ ]” the group can be omitted

`如果八位字节或八位字节组分组在“[]”内，则可以省略该组`

f) Complex APDUs may be built by means of substitutions (sub-structures)

`复杂的APDU可以通过替换（子结构）构建`

g)Exclusive selections of octets or groups of octets are separated by  ^

`八位字节或八位字节组的独占选择由^`

NOTE 1 
The formal PDU example  `正式PDU示例`
AP_PDU = Octet1, OctetGroup1, [Octet2], [Octet3], {[OctGroup2*],OctetGroup3 ^ Octet4}

According to this the following variants are valid on the wire (non exhaustive):

`根据此，以下变体在电线上有效（非详尽）：`

Variant 1: Octet1, OctetGroup1, Octet2, Octet3, OctetGroup2, OctetGroup3
Variant 2: Octet1, OctetGroup1, Octet2, Octet3, OctetGroup2, OctetGroup2, OctetGroup2, OctetGroup3
Variant 3: Octet1, OctetGroup1, OctetGroup2, OctetGroup2, OctetGroup2, OctetGroup3, OctetGroup2
Variant 4: Octet1, OctetGroup1, OctetGroup2, OctetGroup3, OctetGroup2, OctetGroup2, OctetGroup2, OctetGroup2
Variant 5: Octet1, OctetGroup1, Octet3, Octet4

NOTE 2 
The arbitrary order implies that groups of octets are characterised by a special header that is described within the coding rules.

`任意顺序意味着八位字节组的特征在于编码规则中描述的特殊报头。`

NOTE 3
The APDU syntax for RTA-, and RTC-PDU implies that according to the maximum DLSDU an APDU does not exceed 1 440 octets in total.

`RTA-和RTC-PDU的APDU语法意味着根据最大DLSDU，APDU总共不超过1440个八位字节。`

NOTE 4 
The APDU syntax for CL RPC implies that an IO controller supports a minimal ASDU size of 4 096 octets in total and does not exceed 232 P-64 octets in total. The minimal ASDU size is derived from the expected size of configuration, parameter and diagnosis data of an enhanced IO device.

`CL RPC的APDU语法意味着IO控制器支持总共4096个八位字节的最小ASDU大小，并且总共不超过232个P-64八位字节。最小ASDU大小源自增强型IO设备的配置、参数和诊断数据的预期大小。`



##### 4.3.1.3.3 Coding of the field Xid  字段Xid的编码

This field shall be coded as data type Unsigned32. It shall contain a transaction identification chosen by the client to associate requests and responses between a client and a server.
`该字段应编码为数据类型Unsigned32。它应包含客户端选择的事务标识，以关联客户端和服务器之间的请求和响应.`



##### 4.3.1.3.4 Coding of the field DCPDataLength 字段DCPDataLength的编码

This field shall be coded as data type Unsigned16. It shall contain the total length of data followed the DCP-UC-Header or DCP-MC-Header in octets. The maximum length of DCP Data is 1432 bytes.

`该字段应编码为无符号数据类型16。它应包含八位字节中DCP UC头或DCP MC头后面的数据的总长度。DCP数据的最大长度为1432字节。`



#### 4.3.1.4 Coding section of block fields 块字段的编码部分

4.3.1.4.1 General 概述
The block fields are parted into option, suboption, block length, block info, and value. Every block shall assure Unsigned16 alignment. The added padding bytes, with the value zero, shall be counted for the DCPDataLength and shall not be counted for the DCPBlocklength. 805HTable 37 shows the list of available options and 806HTable 38, 807HTable 39, 808HTable 40, 809HTable 41, 810HTable 42, 811HTable 43, and 812HTable 44 shows the list of available suboptions.

`块字段分为选项、子选项、块长度、块信息和值。每个块应确保无符号16对齐。添加的填充字节（值为零）应计入DCPDataLength，而不计入DCPBlocklength。805HT表37显示可用选项列表，806HT表38、807HT表39、808HT表40、809HT表41、810HT表42、811H表43和812H表44显示可用子选项列表。`







