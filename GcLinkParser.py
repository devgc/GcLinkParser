# Tool to parse Link Files and Jump Lists
#
# Matthew Seyer, mseyer@g-cpartners.com
# Copyright 2015 G-C Partners, LLC
#
# Copyright (C) 2015, G-C Partners, LLC <dev@g-cpartners.com>
# G-C Partners licenses this file to you under the Apache License, Version
# 2.0 (the "License"); you may not use this file except in compliance with the
# License.  You may obtain a copy of the License at:
#
#        http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or
# implied.  See the License for the specific language governing
# permissions and limitations under the License.

import logging
log_fmt = '%(module)s:%(funcName)s:%(lineno)d %(message)s'
logging.basicConfig(
    level = logging.DEBUG,
    format=log_fmt
)

import pylnk
import pyfwsi
import pyolecf
import pytz
import argparse
import json
import csv
import struct
import sys
import os
import md5
import datetime
import StringIO
import re

sys.path.append('..\\..\\ElasticHandler\\source')
import elastichandler

_VERSION_ = '1.00'
VERSION = _VERSION_

APPID_STR = '''65009083bfa6a094	(app launched via XPMode)	8/22/2011	Win4n6 List Serv
469e4a7982cea4d4	? (.job)	8/22/2011	Win4n6 List Serv
b0459de4674aab56	(.vmcx)	8/22/2011	Win4n6 List Serv
89b0d939f117f75c	Adobe Acrobat 9 Pro Extended (32-bit)	8/22/2011	Microsoft Windows 7 Forum
26717493b25aa6e1	Adobe Dreamweaver CS5 (32-bit)	8/22/2011	Microsoft Windows 7 Forum
e2a593822e01aed3	Adobe Flash CS5 (32-bit)	8/22/2011	Microsoft Windows 7 Forum
c765823d986857ba	Adobe Illustrator CS5 (32-bit)	8/22/2011	Microsoft Windows 7 Forum
84f066768a22cc4f	Adobe Photoshop CS5 (64-bit)	8/22/2011	Microsoft Windows 7 Forum
44a398496acc926d	Adobe Premiere Pro CS5 (64-bit)	8/22/2011	Microsoft Windows 7 Forum
23646679aaccfae0	Adobe Reader 9.	8/22/2011	Microsoft Windows 7 Forum
23646679aaccfae0	Adobe Reader 9 x64	8/22/2011	Win4n6 List Serv
d5c3931caad5f793	Adobe Soundbooth CS5 (32-bit)	8/22/2011	Microsoft Windows 7 Forum
7e4dca80246863e3	Control Panel (?)	8/22/2011	Win4n6 List Serv
5c450709f7ae4396	Firefox 3.6.13 (32-bit)	8/22/2011	Microsoft Windows 7 Forum
bc03160ee1a59fc1	Foxit PDF Reader 5.4.5	6/7/2013	[ChadTilbury]
28c8b86deab549a1	Internet Explorer 8 / 9 / 10 (32-bit)	8/22/2011	Microsoft Windows 7 Forum
5da8f997fd5f9428	Internet Explorer x64	8/22/2011	Win4n6 List Serv
83b03b46dcd30a0e	iTunes 10	8/22/2011	Win4n6 List Serv
271e609288e1210a	Microsoft Office Access 2010 x86	8/22/2011	Win4n6 List Serv
cdf30b95c55fd785	Microsoft Office Excel 2007	8/22/2011	Win4n6 List Serv
9839aec31243a928	Microsoft Office Excel 2010 x86	8/22/2011	Microsoft Windows 7 Forum
6e855c85de07bc6a	Microsoft Office Excel 2010 x64	6/7/2013	[ChadTilbury]
f0275e8685d95486	Microsoft Office Excel 2013 x86	3/4/2015	Russ Taylor
b8c29862d9f95832	Microsoft Office InfoPath 2010 x86	8/22/2011	Win4n6 List Serv
d64d36b238c843a3	Microsoft Office InfoPath 2010 x86	8/22/2011	Win4n6 List Serv
3094cdb43bf5e9c2	Microsoft Office OneNote 2010 x86	8/22/2011	Win4n6 List Serv
d38adec6953449ba	Microsoft Office OneNote 2010 x64	6/7/2013	[ChadTilbury]
be71009ff8bb02a2	Microsoft Office Outlook x86	8/22/2011	Win4n6 List Serv
5d6f13ed567aa2da	Microsoft Office Outlook 2010 x64	6/7/2013	[ChadTilbury]
f5ac5390b9115fdb	Microsoft Office PowerPoint 2007	8/22/2011	Win4n6 List Serv
9c7cc110ff56d1bd	Microsoft Office PowerPoint 2010 x86	8/22/2011	Microsoft Windows 7 Forum
5f6e7bc0fb699772	Microsoft Office PowerPoint 2010 x64	6/7/2013	[ChadTilbury]
a8c43ef36da523b1	Microsoft Office Word 2003 Pinned and Recent.	8/22/2011	Microsoft Windows 7 Forum
adecfb853d77462a	Microsoft Office Word 2007 Pinned and Recent.	8/22/2011	Microsoft Windows 7 Forum
a7bd71699cd38d1c	Microsoft Office Word 2010 x86	8/22/2011	Microsoft Windows 7 Forum
44a3621b32122d64	Microsoft Office Word 2010 x64	6/7/2013	[ChadTilbury]
a4a5324453625195	Microsoft Office Word 2013 x86	6/9/2014	[BretABennett]
012dc1ea8e34b5a6	Microsoft Paint 6.1	8/22/2011	Win4n6 List Serv
e70d383b15687e37	Notepad++ 5.6.8 (32-bit)	8/22/2011	Microsoft Windows 7 Forum
918e0ecb43d17e23	Notepad (32-bit)	8/22/2011	Microsoft Windows 7 Forum
9b9cdc69c1c24e2b	Notepad (64-bit)	8/22/2011	Microsoft Windows 7 Forum
c7a4093872176c74	Paint Shop Pro Pinned and Recent.	8/22/2011	Microsoft Windows 7 Forum
c71ef2c372d322d7	PGP Desktop 10	8/22/2011	Win4n6 List Serv
431a5b43435cc60b	Python (.pyc)	8/22/2011	Win4n6 List Serv
500b8c1d5302fc9c	Python (.pyw)	8/22/2011	Win4n6 List Serv
1bc392b8e104a00e	Remote Desktop	8/22/2011	Win4n6 List Serv
315e29a36e961336	Roboform 7.8	6/7/2013	[ChadTilbury]
17d3eb086439f0d7	TrueCrypt 7.0a	8/22/2011	Win4n6 List Serv
050620fe75ee0093	VMware Player 3.1.4	8/22/2011	Win4n6 List Serv
8eafbd04ec8631ce	VMware Workstation 9 x64	6/7/2013	[ChadTilbury]
6728dd69a3088f97	Windows Command Processor - cmd.exe (64-bit)	8/22/2011	Microsoft Windows 7 Forum
1b4dd67f29cb1962	Windows Explorer Pinned and Recent.	8/22/2011	Microsoft Windows 7 Forum
f01b4d95cf55d32a	Windows Explorer Windows 8.1.	11/7/2014	Russ Taylor
d7528034b5bd6f28	Windows Live Mail Pinned and Recent.	8/22/2011	Microsoft Windows 7 Forum
b91050d8b077a4e8	Windows Media Center x64	8/22/2011	Win4n6 List Serv
43578521d78096c6	Windows Media Player Classic Home Cinema 1.3 (32-bit)	8/22/2011	Microsoft Windows 7 Forum
74d7f43c1561fc1e	Windows Media Player 12 (32-bit)	8/22/2011	Microsoft Windows 7 Forum
954ea5f70258b502	Windows Script Host - wscript.exe (32-bit)	8/22/2011	Microsoft Windows 7 Forum
9f5c7755804b850a	Windows Script Host - wscript.exe (64-bit)	8/22/2011	Microsoft Windows 7 Forum
b0459de4674aab56	Windows Virtual PC - vmwindow.exe (32- and 64-bit)	8/22/2011	Microsoft Windows 7 Forum
469e4a7982cea4d4	Windows Wordpad	6/7/2013	[ChadTilbury]
290532160612e071	WinRar x64	8/22/2011	Win4n6 List Serv
b74736c2bd8cc8a5	WinZip	8/23/2011	Win4n6 List Serv
e36bfc8972e5ab1d	XPS Viewer	8/22/2011	Win4n6 List Serv
2b53c4ddf69195fc	Zune x64	8/22/2011	Win4n6 List Serv
f0468ce1ae57883d	Adobe Reader 7.1.0	9/8/2011	4n6k Blog
c2d349a0e756411b	Adobe Reader 8.1.2	9/8/2011	4n6k Blog
23646679aaccfae0	Adobe Acrobat 9.4.0	9/8/2011	4n6k Blog
ee462c3b81abb6f6	Adobe Reader X 10.1.0	9/8/2011	4n6k Blog
386a2f6aa7967f36	EyeBrowse 2.7	9/8/2011	4n6k Blog
e31a6a8a7506f733	Image AXS Pro 4.1	9/8/2011	4n6k Blog
b39c5f226977725d	ACDSee Pro 8.1.99	9/8/2011	4n6k Blog
59f56184c796cfd4	ACDSee Photo Manager 10 (Build 219)	9/8/2011	4n6k Blog
8bd5c6433ca967e9	ACDSee Photo Manager 2009 (v11.0 Build 113)	9/8/2011	4n6k Blog
d838aac097abece7	ACDSee Photo Manager 12 (Build 344)	9/8/2011	4n6k Blog
0b3f13480c2785ae	Paint 6.1 (build 7601: SP1)	9/8/2011	4n6k Blog
7cb0735d45243070	CDisplay 1.8.1.0	9/8/2011	4n6k Blog
3594aab44bca414b	Windows Photo Viewer	9/8/2011	4n6k Blog
3edf100b207e2199	digiKam 1.7.0 (KDE 4.4.4)	9/8/2011	4n6k Blog
169b3be0bc43d592	FastPictureViewer Professional 1.6 (Build 211)	9/8/2011	4n6k Blog
e9a39dfba105ea23	FastStone Image Viewer 4.6	9/8/2011	4n6k Blog
0ef606b196796ebb	HP MediaSmart Photo	1/24/2012	[Jimmy_Weg]
edc786643819316c	HoneyView3 #5834	9/8/2011	4n6k Blog
76689ff502a1fd9e	Imagine Image and Animation Viewer 1.0.7	9/8/2011	4n6k Blog
2519133d6d830f7e	IMatch 3.6.0.113	9/8/2011	4n6k Blog
1110d9896dceddb3	imgSeek 0.8.5	9/8/2011	4n6k Blog
c634153e7f5fce9c	IrfanView 3.10 / 4.30	9/8/2011	4n6k Blog
ea83017cdd24374d	IrfanView Thumbnails	9/8/2011	4n6k Blog
3917dd550d7df9a8	Konvertor 4.06 (Build 10)	9/8/2011	4n6k Blog
2fa14c7753239e4c	Paint.NET 2.72 / 3.5.8.4081.24580	9/8/2011	4n6k Blog
d33ecf70f0b74a77	Picasa 2.2.0 (Build 28.08, 0)	9/8/2011	4n6k Blog
0b17d3d0c9ca7e29	Picasa 3.8.0 (Build 117.43, 0)	9/8/2011	4n6k Blog
c5c24a503b1727df	XnView 1.98.2 Small / 1.98.2 Standard	9/8/2011	4n6k Blog
497b42680f564128	Zoner PhotoStudio 13 (Build 7)	9/8/2011	4n6k Blog
e0f7a40340179171	imule 1.4.5 (rev. 749)	9/15/2011	4n6k Blog
76f6f1bd18c19698	aMule 2.2.6	9/15/2011	4n6k Blog
cb5250eaef7e3213	ApexDC++ 1.4.3.957	9/15/2011	4n6k Blog
bfc1d76f16fa778f	Ares (Galaxy) 1.8.4 / 1.9.8 / 2.1.0 / 2.1.7.3041	9/15/2011	4n6k Blog
accca100973ef8dc	Azureus 2.0.8.4	9/15/2011	4n6k Blog
ccb36ff8a8c03b4b	Azureus 2.5.0.4 / Vuze 3.0.5.0	9/15/2011	4n6k Blog
558c5bd9f906860a	BearShare Lite 5.2.5.1	9/15/2011	4n6k Blog
e1d47cb031dafb9f	BearShare 6.0.0.22717 / 8.1.0.70928 / 10.0.0.112380	9/15/2011	4n6k Blog
a31ec95fdd5f350f	BitComet 0.49 / 0.59 / 0.69 / 0.79 / 0.89 / 0.99 / 1.07 / 1.28	9/15/2011	4n6k Blog
bcd7ba75303acbcf	BitLord 1.1	9/15/2011	4n6k Blog
1434d6d62d64857d	BitLord 1.2.0-66	9/15/2011	4n6k Blog
e73d9f534ed5618a	BitSpirit 1.2.0.228 / 2.0 / 2.6.3.168 / 2.7.2.239 / 2.8.0.072 / 3.1.0.077 / 3.6.0.550	9/15/2011	4n6k Blog
c9374251edb4c1a8	BitTornado T-0.3.17	9/15/2011	4n6k Blog
2d61cccb4338dfc8	BitTorrent 5.0.0 / 6.0.0 / 7.2.1 (Build 25548)	9/15/2011	4n6k Blog
ba3a45f7fd2583e1	Blubster 3.1.1	9/15/2011	4n6k Blog
4a7e4f6a181d3d08	broolzShare	9/15/2011	4n6k Blog
f001ea668c0aa916	Cabos 0.8.2	9/15/2011	4n6k Blog
560d789a6a42ad5a	DC++ 0.261 / 0.698 / 0.782 (r2402.1)	9/15/2011	4n6k Blog
4aa2a5710da3efe0	DCSharpHub 2.0.0	9/15/2011	4n6k Blog
2db8e25112ab4453	Deluge 1.3.3	9/15/2011	4n6k Blog
5b186fc4a0b40504	Dtella 1.2.5 (Purdue network only)	9/15/2011	4n6k Blog
2437d4d14b056114	EiskaltDC++ 2.2.3	9/15/2011	4n6k Blog
b3016b8da2077262	eMule 0.50a	9/15/2011	4n6k Blog
cbbe886eca4bfc2d	ExoSee 1.0.0	9/15/2011	4n6k Blog
9ad1ec169bf2da7f	FlylinkDC++ r405 (Build 7358)	9/15/2011	4n6k Blog
4dd48f858b1a6ba7	Free Download Manager 3.0 (Build 852)	9/15/2011	4n6k Blog
f214ca2dd40c59c1	FrostWire 4.20.9	9/15/2011	4n6k Blog
73ce3745a843c0a4	FrostWire 5.1.4	9/15/2011	4n6k Blog
00098b0ef1c84088	fulDC 6.78	9/15/2011	4n6k Blog
e6ea77a1d4553872	Gnucleus 1.8.6.0	9/15/2011	4n6k Blog
ed49e1e6ccdba2f5	GNUnet 0.8.1a	9/15/2011	4n6k Blog
cc4b36fbfb69a757	gtk-gnutella 0.97	9/15/2011	4n6k Blog
a746f9625f7695e8	HeXHub 5.07	9/15/2011	4n6k Blog
223bf0f360c6fea5	I2P 0.8.8 (restartable)	9/15/2011	4n6k Blog
2ff9dc8fb7e11f39	I2P 0.8.8 (no window)	9/15/2011	4n6k Blog
f1a4c04eebef2906	[i2p] Robert 0.0.29 Preferences	9/15/2011	4n6k Blog
c8e4c10e5460b00c	iMesh 6.5.0.16898	9/15/2011	4n6k Blog
f61b65550a84027e	iMesh 11.0.0.112351	9/15/2011	4n6k Blog
d460280b17628695	Java Binary	9/15/2011	4n6k Blog
784182360de0c5b6	Kazaa Lite 1.7.1	9/15/2011	4n6k Blog
a75b276f6e72cf2a	Kazaa Lite Tools K++ 2.7.0	9/15/2011	4n6k Blog
ba132e702c0147ef	KCeasy 0.19-rc1	9/15/2011	4n6k Blog
a8df13a46d66f6b5	Kommute (Calypso) 0.24	9/15/2011	4n6k Blog
c5ef839d8d1c76f4	LimeWire 5.2.13	9/15/2011	4n6k Blog
977a5d147aa093f4	Lphant 3.51	9/15/2011	4n6k Blog
96252daff039437a	Lphant 7.0.0.112351	9/15/2011	4n6k Blog
e76a4ef13fbf2bb1	Manolito 3.1.1	9/15/2011	4n6k Blog
99c15cf3e6d52b61	mldonkey 3.1.0	9/15/2011	4n6k Blog
ff224628f0e8103c	Morpheus 3.0.3.6	9/15/2011	4n6k Blog
792699a1373f1386	Piolet 3.1.1	9/15/2011	4n6k Blog
ca1eb46544793057	RetroShare 0.5.2a (Build 4550)	9/15/2011	4n6k Blog
3cf13d83b0bd3867	RevConnect 0.674p (based on DC++)	9/15/2011	4n6k Blog
05e01ecaf82f7d8e	Scour Exchange 0.0.0.228	9/15/2011	4n6k Blog
5d7b4175afdcc260	Shareaza 2.0.0.0	9/15/2011	4n6k Blog
0b48ce76eda60b97	Shareaza 8.0.0.112300	9/15/2011	4n6k Blog
23f08dab0f6aaf30	SoMud 1.3.3	9/15/2011	4n6k Blog
135df2a440abe9bb	SoulSeek 156c	9/15/2011	4n6k Blog
ecd21b58c2f65a2f	StealthNet 0.8.7.9	9/15/2011	4n6k Blog
5ea2a50c7979fbdc	TrustyFiles 3.1.0.22	9/15/2011	4n6k Blog
cd8cafb0fb6afdab	uTorrent 1.7.7 (Build 8179) / 1.8.5 / 2.0 / 2.21 (Build 25113) / 3.0 (Build 25583)	9/15/2011	4n6k Blog
a75b276f6e72cf2a	WinMX 3.53	9/15/2011	4n6k Blog
490c000889535727	WinMX 4.9.3.0	9/15/2011	4n6k Blog
ac3a63b839ac9d3a	Vuze 4.6.0.4	9/15/2011	4n6k Blog
d28ee773b2cea9b2	3D-FTP 9.0 build 7	9/15/2011	4n6k Blog
cd2acd4089508507	AbsoluteTelnet 9.18 Lite	9/15/2011	4n6k Blog
e6ef42224b845020	ALFTP 5.20.0.4	9/15/2011	4n6k Blog
9e0b3f677a26bbc4	BitKinex 3.2.3	9/15/2011	4n6k Blog
4cdf7858c6673f4b	Bullet Proof FTP 1.26	9/15/2011	4n6k Blog
714b179e552596df	Bullet Proof FTP 2.4.0 (Build 31)	9/15/2011	4n6k Blog
20ef367747c22564	Bullet Proof FTP 2010.75.0.75	9/15/2011	4n6k Blog
044a50e6c87bc012	Classic FTP Plus 2.15	9/15/2011	4n6k Blog
4fceec8e021ac978	CoffeeCup Free FTP 3.5.0.0	9/15/2011	4n6k Blog
8deb27dfa31c5c2a	CoffeeCup Free FTP 4.4 (Build 1904)	9/15/2011	4n6k Blog
49b5edbd92d8cd58	FTP Commander 8.02	9/15/2011	4n6k Blog
6a316aa67a46820b	Core FTP LE 1.3c (Build 1437) / 2.2 (Build 1689)	9/15/2011	4n6k Blog
be4875bb3e0c158f	CrossFTP 1.75a	9/15/2011	4n6k Blog
c04f69101c131440	CuteFTP 5.0 (Build 50.6.10.2)	9/15/2011	4n6k Blog
0a79a7ce3c45d781	CuteFTP 7.1 (Build 06.06.2005.1)	9/15/2011	4n6k Blog
59e86071b87ac1c3	CuteFTP 8.3 (Build 8.3.4.0007)	9/15/2011	4n6k Blog
d8081f151f4bd8a5	CuteFTP 8.3 Lite (Build 8.3.4.0007)	9/15/2011	4n6k Blog
3198e37206f28dc7	CuteFTP 8.3 Professional (Build 8.3.4.0007)	9/15/2011	4n6k Blog
f82607a219af2999	Cyberduck 4.1.2 (Build 8999)	9/15/2011	4n6k Blog
fa7144034d7d083d	Directory Opus 10.0.2.0.4269 (JL tasks supported)	9/15/2011	4n6k Blog
f91fd0c57c4fe449	ExpanDrive 2.1.0	9/15/2011	4n6k Blog
8f852307189803b8	Far Manager 2.0.1807	9/15/2011	4n6k Blog
226400522157fe8b	FileZilla Server 0.9.39 beta	9/15/2011	4n6k Blog
0a1d19afe5a80f80	FileZilla 2.2.32	9/15/2011	4n6k Blog
e107946bb682ce47	FileZilla 3.5.1	9/15/2011	4n6k Blog
b7cb1d1c1991accf	FlashFXP 4.0.0 (Build 1548)	9/15/2011	4n6k Blog
8628e76fd9020e81	Fling File Transfer Plus 2.24	9/15/2011	4n6k Blog
27da120d7e75cf1f	pbFTPClient 6.1	9/15/2011	4n6k Blog
f64de962764b9b0f	FTPRush 1.1.3 / 2.15	9/15/2011	4n6k Blog
10f5a20c21466e85	FTP Voyager 15.2.0.17	9/15/2011	4n6k Blog
7937df3c65790919	FTP Explorer 10.5.19 (Build 001)	9/15/2011	4n6k Blog
9560577fd87cf573	LeechFTP 1.3 (Build 207)	9/15/2011	4n6k Blog
fc999f29bc5c3560	Robo-FTP 3.7.9	9/15/2011	4n6k Blog
c99ddde925d26df3	Robo-FTP 3.7.9 CronMaker	9/15/2011	4n6k Blog
4b632cf2ceceac35	Robo-FTP Server 3.2.5	9/15/2011	4n6k Blog
3a5148bf2288a434	Secure FTP 2.6.1 (Build 20101209.1254)	9/15/2011	4n6k Blog
435a2f986b404eb7	SmartFTP 4.0.1214.0	9/15/2011	4n6k Blog
e42a8e0f4d9b8dcf	Sysax FTP Automation 5.15	9/15/2011	4n6k Blog
b8c13a5dd8c455a2	Titan FTP Server 8.40 (Build 1338)	9/15/2011	4n6k Blog
7904145af324576e	Total Commander 7.56a (Build 16.12.2010)	9/15/2011	4n6k Blog
79370f660ab51725	UploadFTP 2.0.1.0	9/15/2011	4n6k Blog
6a8b377d0f5cb666	WinSCP 2.3.0 (Build 146)	9/15/2011	4n6k Blog
9a3bdae86d5576ee	WinSCP 3.2.1 (Build 174) / 3.8.0 (Build 312)	9/15/2011	4n6k Blog
6bb54d82fa42128d	WinSCP 4.3.4 (Build 1428)	9/15/2011	4n6k Blog
b6267f3fcb700b60	WiseFTP 4.1.0	9/15/2011	4n6k Blog
a581b8002a6eb671	WiseFTP 5.5.9	9/15/2011	4n6k Blog
2544ff74641b639d	WiseFTP 6.1.5	9/15/2011	4n6k Blog
c54b96f328bdc28d	WiseFTP 7.3.0	9/15/2011	4n6k Blog
b223c3ffbc0a7a42	Bersirc 2.2.14	9/15/2011	4n6k Blog
c01d68e40226892b	ClicksAndWhistles 2.7.146	9/15/2011	4n6k Blog
ac8920ed05001800	DMDirc 0.6.5 (Profile store: C:\Users\$user\AppData\Roaming\DMDirc\)	9/15/2011	4n6k Blog
d3530c5294441522	HydraIRC 0.3.165	9/15/2011	4n6k Blog
8904a5fd2d98b546	IceChat 7.70 20101031	9/15/2011	4n6k Blog
6b3a5ce7ad4af9e4	IceChat 9 RC2	9/15/2011	4n6k Blog
fa496fe13dd62edf	KVIrc 3.4.2.1 / 4.0.4	9/15/2011	4n6k Blog
65f7dd884b016ab2	LimeChat 2.39	9/15/2011	4n6k Blog
19ccee0274976da8	mIRC 4.72 / 5.61	9/15/2011	4n6k Blog
ae069d21df1c57df	mIRC 6.35 / 7.19	9/15/2011	4n6k Blog
e30bbea3e1642660	Neebly 1.0.4	9/15/2011	4n6k Blog
54c803dfc87b52ba	Nettalk 6.7.12	9/15/2011	4n6k Blog
dd658a07478b46c2	PIRCH98 1.0.1.1190	9/15/2011	4n6k Blog
6fee01bd55a634fe	Smuxi 0.8.0.0	9/15/2011	4n6k Blog
2a5a615382a84729	X-Chat 2 2.8.6-2	9/15/2011	4n6k Blog
b3965c840bf28ef4	AIM 4.8.2616	9/15/2011	4n6k Blog
01b29f0dc90366bb	AIM 5.9.3857	9/15/2011	4n6k Blog
27ececd8d89b6767	AIM 6.2.14.2 / 6.5.3.12 / 6.9.17.2	9/15/2011	4n6k Blog
0006f647f9488d7a	AIM 7.5.11.9 (custom AppID + JL support)	9/15/2011	4n6k Blog
ca942805559495e9	aMSN 0.98.4	9/15/2011	4n6k Blog
c6f7b5bf1b9675e4	BitWise IM 1.7.3a	9/15/2011	4n6k Blog
fb1f39d1f230480a	Bopup Messenger 5.6.2.9178 (all languages: en;du;fr;ger;rus;es)	9/15/2011	4n6k Blog
dc64de6c91c18300	Brosix Communicator 3.1.3 (Build 110719 nid 1)	9/15/2011	4n6k Blog
f09b920bfb781142	Camfrog 4.0.47 / 5.5.0 / 6.1 (build 146) (JL support)	9/15/2011	4n6k Blog
ebd8c95d87f25154	Carrier 2.5.5	9/15/2011	4n6k Blog
30d23723bdd5d908	Digsby (Build 30140) (JL support)	9/15/2011	4n6k Blog
728008617bc3e34b	eM Client 3.0.10206.0	9/15/2011	4n6k Blog
689319b6547cda85	emesene 2.11.7	9/15/2011	4n6k Blog
454ef7dca3bb16b2	Exodus 0.10.0.0	9/15/2011	4n6k Blog
cca6383a507bac64	Gadu-Gadu 10.5.2.13164	9/15/2011	4n6k Blog
4278d3dc044fc88a	Gaim 1.5.0	9/15/2011	4n6k Blog
777483d3cdac1727	Gajim 0.14.4	9/15/2011	4n6k Blog
6aa18a60024620ae	GCN 2.9.1	9/15/2011	4n6k Blog
3f2cd46691bbee90	GOIM 1.1.0	9/15/2011	4n6k Blog
73c6a317412687c2	Google Talk 1.0.0.104	9/15/2011	4n6k Blog
b0236d03c0627ac4	ICQ 5.1 / ICQLite Build 1068	9/15/2011	4n6k Blog
a5db18f617e28a51	ICQ 6.5 (Build 2024)	9/15/2011	4n6k Blog
2417caa1f2a881d4	ICQ 7.6 (Build 5617)	9/15/2011	4n6k Blog
989d7545c2b2e7b2	IMVU 465.8.0.0	9/15/2011	4n6k Blog
a3e0d98f5653b539	Instantbird 1.0 (20110623121653) (JL support)	9/15/2011	4n6k Blog
bcc705f705d8132b	Instan-t 5.2 (Build 2824)	9/15/2011	4n6k Blog
06059df4b02360af	Kadu 0.10.0 / 0.6.5.5	9/15/2011	4n6k Blog
c312e260e424ae76	Mail.Ru Agent 5.8 (JL support)	9/15/2011	4n6k Blog
22cefa022402327d	Meca Messenger 5.3.0.52	9/15/2011	4n6k Blog
86b804f7a28a3c17	Miranda IM 0.6.8 / 0.7.6 / 0.8.27 / 0.9.9 / 0.9.29 (ANSI + Unicode)	9/15/2011	4n6k Blog
b868d9201b866d96	Microsoft Lync 4.0.7577.0	9/15/2011	4n6k Blog
8c816c711d66a6b5	MSN Messenger 6.2.0137 / 7.0.0820	9/15/2011	4n6k Blog
2d1658d5dc3cbe2d	MySpaceIM 1.0.823.0 Beta	9/15/2011	4n6k Blog
bf9ae1f46bd9c491	Nimbuzz 2.0.0 (rev 6266)	9/15/2011	4n6k Blog
fb7ca8059b8f2123	ooVoo 3.0.7.21	9/15/2011	4n6k Blog
efb08d4e11e21ece	Paltalk Messenger 10.0 (Build 409)	9/15/2011	4n6k Blog
4f24a7b84a7de5a6	Palringo 2.6.3 (r45983)	9/15/2011	4n6k Blog
e93dbdcede8623f2	Pandion 2.6.106	9/15/2011	4n6k Blog
aedd2de3901a77f4	Pidgin 2.0.0 / 2.10.0 / 2.7.3	9/15/2011	4n6k Blog
c5236fd5824c9545	PLAYXPERT 1.0.140.2822	9/15/2011	4n6k Blog
dee18f19c7e3a2ec	PopNote 5.21	9/15/2011	4n6k Blog
1a60b1067913516a	Psi 0.14	9/15/2011	4n6k Blog
e0532b20aa26a0c9	QQ International 1.1 (2042)	9/15/2011	4n6k Blog
3c0022d9de573095	QuteCom 2.2	9/15/2011	4n6k Blog
93b18adf1d948fa3	qutIM 0.2	9/15/2011	4n6k Blog
e0246018261a9ccc	qutIM 0.2.80.0	9/15/2011	4n6k Blog
2aa756186e21b320	RealTimeQuery 3.2	9/15/2011	4n6k Blog
521a29e5d22c13b4	Skype 1.4.0.84 / 2.5.0.154 / 3.8.0.139 / 4.2.0.187 / Skype 5.3.0.120 / 5.5.0.115 / 5.5.32.117	9/15/2011	4n6k Blog
070b52cf73249257	Sococo 1.5.0.2274	9/15/2011	4n6k Blog
d41746b133d17456	Tkabber 0.11.1	9/15/2011	4n6k Blog
c8aa3eaee3d4343d	Trillian 0.74 / 3.1 / 4.2.0.25 / 5.0.0.35 (JL support)	9/15/2011	4n6k Blog
d7d647c92cd5d1e6	uTalk 2.6.4 r47692	9/15/2011	4n6k Blog
36c36598b08891bf	Vovox 2.5.3.4250	9/15/2011	4n6k Blog
884fd37e05659f3a	VZOchat 6.3.5	9/15/2011	4n6k Blog
3461e4d1eb393c9c	WTW 0.8.18.2852 / 0.8.19.2940	9/15/2011	4n6k Blog
f2cb1c38ab948f58	X-Chat 1.8.10 / 2.6.9 / 2.8.9	9/15/2011	4n6k Blog
4e0ac37db19cba15	Xfire 1.138 (Build 44507)	9/15/2011	4n6k Blog
da7e8de5b8273a0f	Yahoo Messenger 5.0.0.1226 / 6.0.0.1922	9/15/2011	4n6k Blog
62dba7fb39bb0adc	Yahoo Messenger 7.5.0.647 / 8.1.0.421 / 9.0.0.2162 / 10.0.0.1270	9/15/2011	4n6k Blog
fb230a9fe81e71a8	Yahoo Messenger 11.0.0.2014-us	9/15/2011	4n6k Blog
b06a975b62567622	Windows Live Messenger 8.5.1235.0517 BETA	9/15/2011	4n6k Blog
bd249197a6faeff2	Windows Live Messenger 2011	9/15/2011	4n6k Blog
d22ad6d9d20e6857	ALLPlayer 4.7	9/8/2011	4n6k Blog
7494a606a9eef18e	Crystal Player 1.98	9/8/2011	4n6k Blog
1cffbe973a437c74	DSPlayer 0.889 Lite	9/8/2011	4n6k Blog
817bb211c92fd254	GOM Player 2.0.12.3375 / 2.1.28.5039	9/8/2011	4n6k Blog
6bc3383cb68a3e37	iTunes 7.6.0.29 / 8.0.0.35	9/8/2011	4n6k Blog
83b03b46dcd30a0e	iTunes 9.0.0.70 / 9.2.1.5 / 10.4.1.10 (begin custom 'Tasks' JL capability)	9/8/2011	4n6k Blog
fe5e840511621941	JetAudio 5.1.9.3018 Basic / 6.2.5.8220 Basic / 7.0.0 Basic / 8.0.16.2000 Basic	9/8/2011	4n6k Blog
a777ad264b54abab	JetVideo 8.0.2.200 Basic	9/8/2011	4n6k Blog
3c93a049a30e25e6	J. River Media Center 16.0.149	9/8/2011	4n6k Blog
4a49906d074a3ad3	Media Go 1.8 (Build 121)	9/8/2011	4n6k Blog
1cf97c38a5881255	MediaPortal 1.1.3	9/8/2011	4n6k Blog
62bff50b969c2575	Quintessential Media Player 5.0 (Build 121) - also usage stats (times used, tracks played, total time used)	9/8/2011	4n6k Blog
b50ee40805bd280f	QuickTime Alternative 1.9.5 (Media Player Classic 6.4.9.1)	9/8/2011	4n6k Blog
ae3f2acd395b622e	QuickTime Player 6.5.1 / 7.0.3 / 7.5.5 (Build 249.13)	9/8/2011	4n6k Blog
7593af37134fd767	RealPlayer 6.0.6.99 / 7 / 8 / 10.5	9/8/2011	4n6k Blog
37392221756de927	RealPlayer SP 12	9/8/2011	4n6k Blog
f92e607f9de02413	RealPlayer 14.0.6.666	9/8/2011	4n6k Blog
6e9d40a4c63bb562	Real Player Alternative 1.25 (Media Player Classic 6.4.8.2 / 6.4.9.0)	9/8/2011	4n6k Blog
c91d08dcfc39a506	SM Player 0.6.9 r3447	9/8/2011	4n6k Blog
e40cb5a291ad1a5b	Songbird 1.9.3 (Build 1959)	9/8/2011	4n6k Blog
4d8bdacf5265a04f	The KMPlayer 2.9.4.1434	9/8/2011	4n6k Blog
4acae695c73a28c7	VLC 0.3.0 / 0.4.6	9/8/2011	4n6k Blog
9fda41b86ddcf1db	VLC 0.5.3 / 0.8.6i / 0.9.7 / 1.1.11	9/8/2011	4n6k Blog
e6ee34ac9913c0a9	VLC 0.6.2	9/8/2011	4n6k Blog
cbeb786f0132005d	VLC 0.7.2	9/8/2011	4n6k Blog
f674c3a77cfe39d0	Winamp 2.95 / 5.1 / 5.621	9/8/2011	4n6k Blog
90e5e8b21d7e7924	Winamp 3.0d (Build 488)	9/8/2011	4n6k Blog
74d7f43c1561fc1e	Windows Media Player 12.0.7601.17514	9/8/2011	4n6k Blog
ed7a5cc3cca8d52a	CCleaner 1.32.345 / 1.41.544 / 2.36.1233 / 3.10.1525	9/8/2011	4n6k Blog
eb7e629258d326a1	WindowWasher 6.6.1.18	9/8/2011	4n6k Blog
ace8715529916d31	40tude Dialog 2.0.15.1 (Beta 38)	9/15/2011	4n6k Blog
cc76755e0f925ce6	AllPicturez 1.2	9/15/2011	4n6k Blog
36f6bc3efe1d99e0	Alt.Binz 0.25.0 (Build 27.09.2007)	9/15/2011	4n6k Blog
d53b52fb65bde78c	Android Newsgroup Downloader 6.2	9/15/2011	4n6k Blog
c845f3a6022d647c	Another File 2.03 (Build 2/7/2004)	9/15/2011	4n6k Blog
780732558f827a42	AutoPix 5.3.3	9/15/2011	4n6k Blog
baea31eacd87186b	BinaryBoy 1.97 (Build 55)	9/15/2011	4n6k Blog
eab25958dbddbaa4	Binary News Reaper 2 (Beta 0.14.7.448)	9/15/2011	4n6k Blog
bf483b423ebbd327	Binary Vortex 5.0	9/15/2011	4n6k Blog
36801066f71b73c5	Binbot 2.0	9/15/2011	4n6k Blog
13eb0e5d9a49eaef	Binjet 3.0.2	9/15/2011	4n6k Blog
8172865a9d5185cb	Binreader 1.0 (Beta 1)	9/15/2011	4n6k Blog
6224453d9701a612	BinTube 3.7.1.0 (requires VLC 10.5!)	9/15/2011	4n6k Blog
cf6379a9a987366e	Digibin 1.31	9/15/2011	4n6k Blog
43886ba3395acdcc	Easy Post 3.0	9/15/2011	4n6k Blog
0cfab0ec14b6f953	Express NewsPictures 2.41 (Build 08.05.07.0)	9/15/2011	4n6k Blog
7526de4a8b5914d9	Forte Agent 6.00 (Build 32.1186)	9/15/2011	4n6k Blog
c02baf50d02056fc	FotoVac 1.0	9/15/2011	4n6k Blog
3ed70ef3495535f7	Gravity 3.0.4	9/15/2011	4n6k Blog
86781fe8437db23e	Messenger Pro 2.66.6.3353	9/15/2011	4n6k Blog
f920768fe275f7f4	Grabit 1.5.3 Beta (Build 909) / 1.6.2 (Build 940) / 1.7.2 Beta 4 (Build 997)	9/15/2011	4n6k Blog
9f03ae476ad461fa	GroupsAloud 1.0	9/15/2011	4n6k Blog
d0261ed6e16b200b	News File Grabber 4.6.0.4	9/15/2011	4n6k Blog
8211531a7918b389	Newsbin Pro 6.00 (Build 1019) (JL support)	9/15/2011	4n6k Blog
d1fc019238236806	Newsgroup Commander Pro 9.05	9/15/2011	4n6k Blog
186b5ccada1d986b	NewsGrabber 3.0.36	9/15/2011	4n6k Blog
4d72cfa1d0a67418	Newsgroup Image Collector	9/15/2011	4n6k Blog
92f1d5db021cd876	NewsLeecher 4.0 / 5.0 Beta 6	9/15/2011	4n6k Blog
d7666c416cba240c	NewsMan Pro 3.0.5.2	9/15/2011	4n6k Blog
7b2b4f995b54387d	News Reactor 20100224.16	9/15/2011	4n6k Blog
cb984e3bc7faf234	NewsRover 17.0 (Rev.0)	9/15/2011	4n6k Blog
c98ab5ccf25dda79	NewsShark 2.0	9/15/2011	4n6k Blog
dba909a61476ccec	NewsWolf 1.41	9/15/2011	4n6k Blog
2b164f512891ae37	NewsWolf NSListGen	9/15/2011	4n6k Blog
cb1d97aca3fb7e6b	Newz Crawler 1.9.0 (Build 4100)	9/15/2011	4n6k Blog
3be7b307dfccb58f	NiouzeFire 0.8.7.0	9/15/2011	4n6k Blog
de76415e0060ce13	Noworyta News Reader 2.9	9/15/2011	4n6k Blog
cd40ead0b1eb15ab	NNTPGrab 0.6.2	9/15/2011	4n6k Blog
d5c02fc7afbb3fd4	NNTPGrab 0.6.2 Server	9/15/2011	4n6k Blog
a4def57ee99d77e9	Nomad News 1.43	9/15/2011	4n6k Blog
3f97341a65bac63a	Ozum 6.07 (Build 6070)	9/15/2011	4n6k Blog
bfe841f4d35c92b1	QuadSucker/News 5.0	9/15/2011	4n6k Blog
d3c5cf21e86b28af	SeaMonkey 2.3.3	9/15/2011	4n6k Blog
7a7c60efd66817a2	Spotnet 1.7.4	9/15/2011	4n6k Blog
eb3300e672136bc7	Stream Reactor 1.0 Beta 9 (uses VLC!)	9/15/2011	4n6k Blog
3168cc975b354a01	Slypheed 3.1.2 (Build 1120)	9/15/2011	4n6k Blog
776beb1fcfc6dfa5	Thunderbird 1.0.6 (20050716) / 3.0.2	9/15/2011	4n6k Blog
03d877ec11607fe4	Thunderbird 6.0.2	9/15/2011	4n6k Blog
7192f2de78fd9e96	TIFNY 5.0.3	9/15/2011	4n6k Blog
9dacebaa9ac8ca4e	TLNews Newsreader 2.2.0 (Build 2430)	9/15/2011	4n6k Blog
7fd04185af357bd5	UltraLeeacher 1.7.0.2969 / 1.8 Beta (Build 3490)	9/15/2011	4n6k Blog
aa11f575087b3bdc	Unzbin 2.6.8	9/15/2011	4n6k Blog
d7db75db9cdd7c5d	Xnews 5.04.25	9/15/2011	4n6k Blog
3dc02b55e44d6697	7-Zip 3.13 / 4.20	9/8/2011	4n6k Blog
4975d6798a8bdf66	7-Zip 4.65 / 9.20	9/8/2011	4n6k Blog
4b6925efc53a3c08	BCWipe 5.02.2 Task Manager 3.02.3	9/8/2011	4n6k Blog
23709f6439b9f03d	Hex Editor Neo 5.14	6/7/2013	[ChadTilbury]
e57cfc995bdc1d98	Snagit 11	6/7/2013	[ChadTilbury]
337ed59af273c758	Sticky Notes	9/8/2011	4n6k Blog
290532160612e071	WinRAR 2.90 / 3.60 / 4.01	9/8/2011	4n6k Blog
c9950c443027c765	WinZip 9.0 SR-1 (6224) / 10.0 (6667)	9/8/2011	4n6k Blog
b74736c2bd8cc8a5	WinZip 15.5 (9468)	9/8/2011	4n6k Blog
bc0c37e84e063727	Windows Command Processor - cmd.exe (32-bit)	9/8/2011	4n6k Blog'''

TIMEZONES = [
    "Africa/Abidjan",
    "Africa/Accra",
    "Africa/Addis_Ababa",
    "Africa/Algiers",
    "Africa/Asmara",
    "Africa/Asmera",
    "Africa/Bamako",
    "Africa/Bangui",
    "Africa/Banjul",
    "Africa/Bissau",
    "Africa/Blantyre",
    "Africa/Brazzaville",
    "Africa/Bujumbura",
    "Africa/Cairo",
    "Africa/Casablanca",
    "Africa/Ceuta",
    "Africa/Conakry",
    "Africa/Dakar",
    "Africa/Dar_es_Salaam",
    "Africa/Djibouti",
    "Africa/Douala",
    "Africa/El_Aaiun",
    "Africa/Freetown",
    "Africa/Gaborone",
    "Africa/Harare",
    "Africa/Johannesburg",
    "Africa/Juba",
    "Africa/Kampala",
    "Africa/Khartoum",
    "Africa/Kigali",
    "Africa/Kinshasa",
    "Africa/Lagos",
    "Africa/Libreville",
    "Africa/Lome",
    "Africa/Luanda",
    "Africa/Lubumbashi",
    "Africa/Lusaka",
    "Africa/Malabo",
    "Africa/Maputo",
    "Africa/Maseru",
    "Africa/Mbabane",
    "Africa/Mogadishu",
    "Africa/Monrovia",
    "Africa/Nairobi",
    "Africa/Ndjamena",
    "Africa/Niamey",
    "Africa/Nouakchott",
    "Africa/Ouagadougou",
    "Africa/Porto-Novo",
    "Africa/Sao_Tome",
    "Africa/Timbuktu",
    "Africa/Tripoli",
    "Africa/Tunis",
    "Africa/Windhoek",
    "America/Adak",
    "America/Anchorage",
    "America/Anguilla",
    "America/Antigua",
    "America/Araguaina",
    "America/Argentina/Buenos_Aires",
    "America/Argentina/Catamarca",
    "America/Argentina/ComodRivadavia",
    "America/Argentina/Cordoba",
    "America/Argentina/Jujuy",
    "America/Argentina/La_Rioja",
    "America/Argentina/Mendoza",
    "America/Argentina/Rio_Gallegos",
    "America/Argentina/Salta",
    "America/Argentina/San_Juan",
    "America/Argentina/San_Luis",
    "America/Argentina/Tucuman",
    "America/Argentina/Ushuaia",
    "America/Aruba",
    "America/Asuncion",
    "America/Atikokan",
    "America/Atka",
    "America/Bahia",
    "America/Bahia_Banderas",
    "America/Barbados",
    "America/Belem",
    "America/Belize",
    "America/Blanc-Sablon",
    "America/Boa_Vista",
    "America/Bogota",
    "America/Boise",
    "America/Buenos_Aires",
    "America/Cambridge_Bay",
    "America/Campo_Grande",
    "America/Cancun",
    "America/Caracas",
    "America/Catamarca",
    "America/Cayenne",
    "America/Cayman",
    "America/Chicago",
    "America/Chihuahua",
    "America/Coral_Harbour",
    "America/Cordoba",
    "America/Costa_Rica",
    "America/Creston",
    "America/Cuiaba",
    "America/Curacao",
    "America/Danmarkshavn",
    "America/Dawson",
    "America/Dawson_Creek",
    "America/Denver",
    "America/Detroit",
    "America/Dominica",
    "America/Edmonton",
    "America/Eirunepe",
    "America/El_Salvador",
    "America/Ensenada",
    "America/Fort_Wayne",
    "America/Fortaleza",
    "America/Glace_Bay",
    "America/Godthab",
    "America/Goose_Bay",
    "America/Grand_Turk",
    "America/Grenada",
    "America/Guadeloupe",
    "America/Guatemala",
    "America/Guayaquil",
    "America/Guyana",
    "America/Halifax",
    "America/Havana",
    "America/Hermosillo",
    "America/Indiana/Indianapolis",
    "America/Indiana/Knox",
    "America/Indiana/Marengo",
    "America/Indiana/Petersburg",
    "America/Indiana/Tell_City",
    "America/Indiana/Vevay",
    "America/Indiana/Vincennes",
    "America/Indiana/Winamac",
    "America/Indianapolis",
    "America/Inuvik",
    "America/Iqaluit",
    "America/Jamaica",
    "America/Jujuy",
    "America/Juneau",
    "America/Kentucky/Louisville",
    "America/Kentucky/Monticello",
    "America/Knox_IN",
    "America/Kralendijk",
    "America/La_Paz",
    "America/Lima",
    "America/Los_Angeles",
    "America/Louisville",
    "America/Lower_Princes",
    "America/Maceio",
    "America/Managua",
    "America/Manaus",
    "America/Marigot",
    "America/Martinique",
    "America/Matamoros",
    "America/Mazatlan",
    "America/Mendoza",
    "America/Menominee",
    "America/Merida",
    "America/Metlakatla",
    "America/Mexico_City",
    "America/Miquelon",
    "America/Moncton",
    "America/Monterrey",
    "America/Montevideo",
    "America/Montreal",
    "America/Montserrat",
    "America/Nassau",
    "America/New_York",
    "America/Nipigon",
    "America/Nome",
    "America/Noronha",
    "America/North_Dakota/Beulah",
    "America/North_Dakota/Center",
    "America/North_Dakota/New_Salem",
    "America/Ojinaga",
    "America/Panama",
    "America/Pangnirtung",
    "America/Paramaribo",
    "America/Phoenix",
    "America/Port-au-Prince",
    "America/Port_of_Spain",
    "America/Porto_Acre",
    "America/Porto_Velho",
    "America/Puerto_Rico",
    "America/Rainy_River",
    "America/Rankin_Inlet",
    "America/Recife",
    "America/Regina",
    "America/Resolute",
    "America/Rio_Branco",
    "America/Rosario",
    "America/Santa_Isabel",
    "America/Santarem",
    "America/Santiago",
    "America/Santo_Domingo",
    "America/Sao_Paulo",
    "America/Scoresbysund",
    "America/Shiprock",
    "America/Sitka",
    "America/St_Barthelemy",
    "America/St_Johns",
    "America/St_Kitts",
    "America/St_Lucia",
    "America/St_Thomas",
    "America/St_Vincent",
    "America/Swift_Current",
    "America/Tegucigalpa",
    "America/Thule",
    "America/Thunder_Bay",
    "America/Tijuana",
    "America/Toronto",
    "America/Tortola",
    "America/Vancouver",
    "America/Virgin",
    "America/Whitehorse",
    "America/Winnipeg",
    "America/Yakutat",
    "America/Yellowknife",
    "Antarctica/Casey",
    "Antarctica/Davis",
    "Antarctica/DumontDUrville",
    "Antarctica/Macquarie",
    "Antarctica/Mawson",
    "Antarctica/McMurdo",
    "Antarctica/Palmer",
    "Antarctica/Rothera",
    "Antarctica/South_Pole",
    "Antarctica/Syowa",
    "Antarctica/Troll",
    "Antarctica/Vostok",
    "Arctic/Longyearbyen",
    "Asia/Aden",
    "Asia/Almaty",
    "Asia/Amman",
    "Asia/Anadyr",
    "Asia/Aqtau",
    "Asia/Aqtobe",
    "Asia/Ashgabat",
    "Asia/Ashkhabad",
    "Asia/Baghdad",
    "Asia/Bahrain",
    "Asia/Baku",
    "Asia/Bangkok",
    "Asia/Beirut",
    "Asia/Bishkek",
    "Asia/Brunei",
    "Asia/Calcutta",
    "Asia/Chita",
    "Asia/Choibalsan",
    "Asia/Chongqing",
    "Asia/Chungking",
    "Asia/Colombo",
    "Asia/Dacca",
    "Asia/Damascus",
    "Asia/Dhaka",
    "Asia/Dili",
    "Asia/Dubai",
    "Asia/Dushanbe",
    "Asia/Gaza",
    "Asia/Harbin",
    "Asia/Hebron",
    "Asia/Ho_Chi_Minh",
    "Asia/Hong_Kong",
    "Asia/Hovd",
    "Asia/Irkutsk",
    "Asia/Istanbul",
    "Asia/Jakarta",
    "Asia/Jayapura",
    "Asia/Jerusalem",
    "Asia/Kabul",
    "Asia/Kamchatka",
    "Asia/Karachi",
    "Asia/Kashgar",
    "Asia/Kathmandu",
    "Asia/Katmandu",
    "Asia/Khandyga",
    "Asia/Kolkata",
    "Asia/Krasnoyarsk",
    "Asia/Kuala_Lumpur",
    "Asia/Kuching",
    "Asia/Kuwait",
    "Asia/Macao",
    "Asia/Macau",
    "Asia/Magadan",
    "Asia/Makassar",
    "Asia/Manila",
    "Asia/Muscat",
    "Asia/Nicosia",
    "Asia/Novokuznetsk",
    "Asia/Novosibirsk",
    "Asia/Omsk",
    "Asia/Oral",
    "Asia/Phnom_Penh",
    "Asia/Pontianak",
    "Asia/Pyongyang",
    "Asia/Qatar",
    "Asia/Qyzylorda",
    "Asia/Rangoon",
    "Asia/Riyadh",
    "Asia/Saigon",
    "Asia/Sakhalin",
    "Asia/Samarkand",
    "Asia/Seoul",
    "Asia/Shanghai",
    "Asia/Singapore",
    "Asia/Srednekolymsk",
    "Asia/Taipei",
    "Asia/Tashkent",
    "Asia/Tbilisi",
    "Asia/Tehran",
    "Asia/Tel_Aviv",
    "Asia/Thimbu",
    "Asia/Thimphu",
    "Asia/Tokyo",
    "Asia/Ujung_Pandang",
    "Asia/Ulaanbaatar",
    "Asia/Ulan_Bator",
    "Asia/Urumqi",
    "Asia/Ust-Nera",
    "Asia/Vientiane",
    "Asia/Vladivostok",
    "Asia/Yakutsk",
    "Asia/Yekaterinburg",
    "Asia/Yerevan",
    "Atlantic/Azores",
    "Atlantic/Bermuda",
    "Atlantic/Canary",
    "Atlantic/Cape_Verde",
    "Atlantic/Faeroe",
    "Atlantic/Faroe",
    "Atlantic/Jan_Mayen",
    "Atlantic/Madeira",
    "Atlantic/Reykjavik",
    "Atlantic/South_Georgia",
    "Atlantic/St_Helena",
    "Atlantic/Stanley",
    "Australia/ACT",
    "Australia/Adelaide",
    "Australia/Brisbane",
    "Australia/Broken_Hill",
    "Australia/Canberra",
    "Australia/Currie",
    "Australia/Darwin",
    "Australia/Eucla",
    "Australia/Hobart",
    "Australia/LHI",
    "Australia/Lindeman",
    "Australia/Lord_Howe",
    "Australia/Melbourne",
    "Australia/NSW",
    "Australia/North",
    "Australia/Perth",
    "Australia/Queensland",
    "Australia/South",
    "Australia/Sydney",
    "Australia/Tasmania",
    "Australia/Victoria",
    "Australia/West",
    "Australia/Yancowinna",
    "Brazil/Acre",
    "Brazil/DeNoronha",
    "Brazil/East",
    "Brazil/West",
    "CET",
    "CST6CDT",
    "Canada/Atlantic",
    "Canada/Central",
    "Canada/East-Saskatchewan",
    "Canada/Eastern",
    "Canada/Mountain",
    "Canada/Newfoundland",
    "Canada/Pacific",
    "Canada/Saskatchewan",
    "Canada/Yukon",
    "Chile/Continental",
    "Chile/EasterIsland",
    "Cuba",
    "EET",
    "EST",
    "EST5EDT",
    "Egypt",
    "Eire",
    "Etc/GMT",
    "Etc/GMT+0",
    "Etc/GMT+1",
    "Etc/GMT+10",
    "Etc/GMT+11",
    "Etc/GMT+12",
    "Etc/GMT+2",
    "Etc/GMT+3",
    "Etc/GMT+4",
    "Etc/GMT+5",
    "Etc/GMT+6",
    "Etc/GMT+7",
    "Etc/GMT+8",
    "Etc/GMT+9",
    "Etc/GMT-0",
    "Etc/GMT-1",
    "Etc/GMT-10",
    "Etc/GMT-11",
    "Etc/GMT-12",
    "Etc/GMT-13",
    "Etc/GMT-14",
    "Etc/GMT-2",
    "Etc/GMT-3",
    "Etc/GMT-4",
    "Etc/GMT-5",
    "Etc/GMT-6",
    "Etc/GMT-7",
    "Etc/GMT-8",
    "Etc/GMT-9",
    "Etc/GMT0",
    "Etc/Greenwich",
    "Etc/UCT",
    "Etc/UTC",
    "Etc/Universal",
    "Etc/Zulu",
    "Europe/Amsterdam",
    "Europe/Andorra",
    "Europe/Athens",
    "Europe/Belfast",
    "Europe/Belgrade",
    "Europe/Berlin",
    "Europe/Bratislava",
    "Europe/Brussels",
    "Europe/Bucharest",
    "Europe/Budapest",
    "Europe/Busingen",
    "Europe/Chisinau",
    "Europe/Copenhagen",
    "Europe/Dublin",
    "Europe/Gibraltar",
    "Europe/Guernsey",
    "Europe/Helsinki",
    "Europe/Isle_of_Man",
    "Europe/Istanbul",
    "Europe/Jersey",
    "Europe/Kaliningrad",
    "Europe/Kiev",
    "Europe/Lisbon",
    "Europe/Ljubljana",
    "Europe/London",
    "Europe/Luxembourg",
    "Europe/Madrid",
    "Europe/Malta",
    "Europe/Mariehamn",
    "Europe/Minsk",
    "Europe/Monaco",
    "Europe/Moscow",
    "Europe/Nicosia",
    "Europe/Oslo",
    "Europe/Paris",
    "Europe/Podgorica",
    "Europe/Prague",
    "Europe/Riga",
    "Europe/Rome",
    "Europe/Samara",
    "Europe/San_Marino",
    "Europe/Sarajevo",
    "Europe/Simferopol",
    "Europe/Skopje",
    "Europe/Sofia",
    "Europe/Stockholm",
    "Europe/Tallinn",
    "Europe/Tirane",
    "Europe/Tiraspol",
    "Europe/Uzhgorod",
    "Europe/Vaduz",
    "Europe/Vatican",
    "Europe/Vienna",
    "Europe/Vilnius",
    "Europe/Volgograd",
    "Europe/Warsaw",
    "Europe/Zagreb",
    "Europe/Zaporozhye",
    "Europe/Zurich",
    "GB",
    "GB-Eire",
    "GMT",
    "GMT+0",
    "GMT-0",
    "GMT0",
    "Greenwich",
    "HST",
    "Hongkong",
    "Iceland",
    "Indian/Antananarivo",
    "Indian/Chagos",
    "Indian/Christmas",
    "Indian/Cocos",
    "Indian/Comoro",
    "Indian/Kerguelen",
    "Indian/Mahe",
    "Indian/Maldives",
    "Indian/Mauritius",
    "Indian/Mayotte",
    "Indian/Reunion",
    "Iran",
    "Israel",
    "Jamaica",
    "Japan",
    "Kwajalein",
    "Libya",
    "MET",
    "MST",
    "MST7MDT",
    "Mexico/BajaNorte",
    "Mexico/BajaSur",
    "Mexico/General",
    "NZ",
    "NZ-CHAT",
    "Navajo",
    "PRC",
    "PST8PDT",
    "Pacific/Apia",
    "Pacific/Auckland",
    "Pacific/Chatham",
    "Pacific/Chuuk",
    "Pacific/Easter",
    "Pacific/Efate",
    "Pacific/Enderbury",
    "Pacific/Fakaofo",
    "Pacific/Fiji",
    "Pacific/Funafuti",
    "Pacific/Galapagos",
    "Pacific/Gambier",
    "Pacific/Guadalcanal",
    "Pacific/Guam",
    "Pacific/Honolulu",
    "Pacific/Johnston",
    "Pacific/Kiritimati",
    "Pacific/Kosrae",
    "Pacific/Kwajalein",
    "Pacific/Majuro",
    "Pacific/Marquesas",
    "Pacific/Midway",
    "Pacific/Nauru",
    "Pacific/Niue",
    "Pacific/Norfolk",
    "Pacific/Noumea",
    "Pacific/Pago_Pago",
    "Pacific/Palau",
    "Pacific/Pitcairn",
    "Pacific/Pohnpei",
    "Pacific/Ponape",
    "Pacific/Port_Moresby",
    "Pacific/Rarotonga",
    "Pacific/Saipan",
    "Pacific/Samoa",
    "Pacific/Tahiti",
    "Pacific/Tarawa",
    "Pacific/Tongatapu",
    "Pacific/Truk",
    "Pacific/Wake",
    "Pacific/Wallis",
    "Pacific/Yap",
    "Poland",
    "Portugal",
    "ROC",
    "ROK",
    "Singapore",
    "Turkey",
    "UCT",
    "US/Alaska",
    "US/Aleutian",
    "US/Arizona",
    "US/Central",
    "US/East-Indiana",
    "US/Eastern",
    "US/Hawaii",
    "US/Indiana-Starke",
    "US/Michigan",
    "US/Mountain",
    "US/Pacific",
    "US/Pacific-New",
    "US/Samoa",
    "UTC",
    "Universal",
    "W-SU",
    "WET",
    "Zulu"
]

COLUMNS = [
    "DriveType",
    "DriveSerialNumber",
    "VolumeLabel",
    "LocalPath",
    "NetworkPath",
    "RelativePath",
    "AppIdCode",
    "AppIdName",
    "BaseName",
    "FileExt",
    "DLE_EntryNum",
    "DLE_BirthDroidVolId",
    "DLE_BirthDroidFileId",
    "DLE_DroidVolId",
    "DLE_DroidFileId",
    "DLE_LastModificationDatetime",
    "DLE_Hostname",
    "DLE_Path",
    "CreationDateTime",
    "ModificationDateTime",
    "AccessDateTime",
    "Codepage",
    "CmdArgs",
    "Description",
    "EnvVarLoc",
    "Flags",
    "FileSize",
    "IconLoc",
    "ParentLongName",
    "ParentEntryNum",
    "ParentSeqNum",
    "WorkingDir",
    "Source",
    "LnkTrgData"
]

DRIVE_TYPES = {
    0:'DRIVE_UNKNOWN;',
    1:'DRIVE_NO_ROOT_DIR;',
    2:'DRIVE_REMOVABLE;',
    3:'DRIVE_FIXED;',
    4:'DRIVE_REMOTE;',
    5:'DRIVE_CDROM;',
    6:'DRIVE_RAMDISK;'
}

FILE_ATTRIBUTE_FLAGS = {
    0x00000001:'Read-Only;',
    0x00000002:'Hidden;',
    0x00000004:'System;',
    0x00000020:'Archive;',
    0x00000040:'Device;',
    0x00000080:'Normal;',
    0x00000100:'Temporary;',
    0x00000200:'Sparse File;',
    0x00000400:'Reparse Point;',
    0x00000800:'Compressed;',
    0x00001000:'Offline;',
    0x00002000:'Not Content Indexed;',
    0x00004000:'Encrypted;',
    0x10000000:'Directory;',
    0x20000000:'Index View;'
}

LINKMAPPING = {
    "mappings": {
        "gc_link_file": {
            "properties": {
                "index_timestamp":{
                    "type": "date",
                    "format": "MM/dd/yyyy HH:mm:ss.SSSSSS||MM/dd/yyyy HH:mm:ss||yyyy-MM-dd HH:mm:ss"
                },
                "ExternalTag":{
                    "type": "string"
                },
                "BaseName":{
                    "type": "string",
                    "index": "not_analyzed"
                },
                "FileExt":{
                    "type": "string",
                    "index": "not_analyzed"
                },
                "EntryNames":{
                    "type": "string",
                    "index": "not_analyzed"
                },
                "EntryReferences":{
                    "type": "string",
                    "index": "not_analyzed"
                },
                "index_timestamp":{
                    "type": "date",
                    "format": "MM/dd/yyyy HH:mm:ss.SSSSSS||MM/dd/yyyy HH:mm:ss||yyyy-MM-DD HH:mm:ss"
                },
                "IconLoc": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "Description": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "RelativePath": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "AccessDateTime": {
                    "type": "date",
                    "format": "MM/dd/yyyy HH:mm:ss.SSSSSS||MM/dd/yyyy HH:mm:ss||yyyy-MM-DD HH:mm:ss"
                },
                "FileName": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "ModificationDateTime": {
                    "type": "date",
                    "format": "MM/dd/yyyy HH:mm:ss.SSSSSS||MM/dd/yyyy HH:mm:ss||yyyy-MM-DD HH:mm:ss"
                },
                "Flags": {
                    "type": "string"
                },
                "VolumeLabel": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "Source": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "DriveSerialNumber": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "NetworkPath": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "DriveType": {
                    "type": "string"
                },
                "WorkingDir": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "CreationDateTime": {
                    "type": "date",
                    "format": "MM/dd/yyyy HH:mm:ss.SSSSSS||MM/dd/yyyy HH:mm:ss||yyyy-MM-DD HH:mm:ss"
                },
                "CmdArgs": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "Codepage": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "LnkTrgData": {
                    "properties": {
                        "NetworkLocation": {
                            "properties": {
                                "Type": {
                                    "type": "string",
                                    "index": "not_analyzed"
                                },
                                "Description": {
                                    "type": "string",
                                    "index": "not_analyzed"
                                },
                                "Location": {
                                    "type": "string",
                                    "index": "not_analyzed"
                                }
                            }
                        },
                        "ParentEntryNum": {
                            "type": "long"
                        },
                        "Volume": {
                            "properties": {
                                "ShellItemTypeStr": {
									"type": "string"
								},
								"Type": {
									"index": "not_analyzed",
									"type": "string"
								},
								"Identifier": {
									"index": "not_analyzed",
									"type": "string"
								},
								"ShellFolderId": {
									"index": "not_analyzed",
									"type": "string"
								},
								"ExtentionBlockCount": {
									"type": "long"
								},
								"ShellItemTypeHex": {
									"type": "string"
								},
								"Name": {
									"index": "not_analyzed",
									"type": "string"
								}
                            }
                        },
                        "ParentSeqNum": {
                            "type": "long"
                        },
                        "DistinctTypes": {
                            "type": "string"
                        },
                        "ItemCount": {
                            "type": "long"
                        },
                        "RootFolder": {
                            "properties": {
                                "ShellItemTypeStr": {
									"type": "string"
								},
								"Type": {
									"type": "string"
								},
								"ShellFolderId": {
									"index": "not_analyzed",
									"type": "string"
								},
								"ExtentionBlockCount": {
									"type": "long"
								},
								"ShellItemTypeHex": {
									"type": "string"
								}
                            }
                        },
                        "FileEntries": {
							"properties": {
								"ExtentionBlocks": {
									"properties": {
										"SeqNum": {
											"type": "long"
										},
										"LocalizedName": {
											"type": "string",
											"index": "not_analyzed"
										},
										"ModificationTime": {
											"type": "date",
											"format": "MM/dd/yyyy HH:mm:ss.SSSSSS||MM/dd/yyyy HH:mm:ss||yyyy-MM-DD HH:mm:ss"
										},
										"Name": {
											"type": "string",
											"index": "not_analyzed"
										},
										"AccessTime": {
											"type": "date",
											"format": "MM/dd/yyyy HH:mm:ss.SSSSSS||MM/dd/yyyy HH:mm:ss||yyyy-MM-DD HH:mm:ss"
										},
										"Type": {
											"type": "string",
											"index": "not_analyzed"
										},
										"LongName": {
											"type": "string",
											"index": "not_analyzed"
										},
										"DataSize": {
											"type": "long"
										},
										"Signature": {
											"type": "string",
											"index": "not_analyzed"
										},
										"CreationTime": {
											"type": "date",
											"format": "MM/dd/yyyy HH:mm:ss.SSSSSS||MM/dd/yyyy HH:mm:ss||yyyy-MM-DD HH:mm:ss"
										},
										"FileReferenceInt": {
											"type": "long"
										},
										"EntryNum": {
											"type": "long"
										},
										"FileSize": {
											"type": "long"
										}
									}
								}
                            }
                        },
                        "ParentLongName": {
                            "type": "string",
                            "index": "not_analyzed"
                        }
                    }
                },
                "EnvVarLoc": {
                    "type": "string",
                    "index": "not_analyzed"
                },
                "FileSize": {
                    "type": "long"
                }
            }
        }
    }
}

def GetOptions():
    '''Get needed options for processesing'''
    #Options:
    #evidence_name,index,report,config
    
    usage = '''GcLinkParser v'''+_VERSION_+''' [Copywrite G-C Partners, LLC 2015,2016]

EXAMPLES:
========================================================================
List Supported Timezones
GcLinkParser.exe --tzlist
------------------------------------------------------------------------
JSON Output
GcLinkParser.exe -f LINKFILE --json
------------------------------------------------------------------------
CSV Output
GcLinkParser.exe -f LINKFILE --txt
------------------------------------------------------------------------
Send records to Elasticsearch
GcLinkParser.exe -f LINKFILE --eshost "ELASTIC_IP" --index lnkfiles
------------------------------------------------------------------------
Get Filelist from dir and format txt
dir /b /s /a *.lnk | GcLinkParser.exe --pipe --txt

NOTES:
========================================================================
The AppId is enumerated from the list found at
http://forensicswiki.org/wiki/List_of_Jump_List_IDs
that was last modified as of 4 March 2015, at 11:07.

You can create a custom AppId list and stick it in the cwd of this tool.
Name the file 'AppIdList.txt' and should be formated as 16HEXID\\tAPP_NAME

'''
    options = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description=(usage)
    )
    
    options.add_argument(
        '-f','--file',
        dest='file_name',
        action="store",
        type=str,
        default=None,
        help='lnk filename'
    )
    
    options.add_argument(
        '--jmp',
        dest='jmp_flag',
        action="store_true",
        default=False,
        help='Parse files as jumplists'
    )
    
    options.add_argument(
        '--pipe',
        dest='pipe_flag',
        action="store_true",
        default=False,
        help='get filelist from pipe (dir /b /s /a *.lnk)'
    )
    
    options.add_argument(
        '--timeformat',
        dest='timeformat',
        action="store",
        type=str,
        default='{0.month:02d}/{0.day:02d}/{0.year:04d} {0.hour:02d}:{0.minute:02d}:{0.second:02d}.{0.resolution.microseconds:06d}',
        help='datetime format'
    )
    
    options.add_argument(
        '--timezone',
        dest='timezone',
        action="store",
        type=str,
        default='UTC',
        help='output timezone'
    )
    
    options.add_argument(
        '--listtz',
        dest='listtz_flag',
        action="store_true",
        default=False,
        help='list all supported timezone options'
    )
    
    options.add_argument(
        '--txt',
        dest='csv_flag',
        action="store_true",
        default=False,
        help='output to text file (default delimiter is \\t) [Recommended not to use \',\' as delimiter]'
    )
    
    options.add_argument(
        '--delimiter',
        dest='delimiter',
        action="store",
        type=str,
        default="\t",
        help='csv delimiter'
    )
    
    options.add_argument(
        '--json',
        dest='json_flag',
        action="store_true",
        default=False,
        help='json output'
    )
    
    options.add_argument(
        '--eshost',
        dest='eshost',
        action="store",
        type=str,
        default=None,
        help='Elastic host'
    )
    
    options.add_argument(
        '--index',
        dest='index',
        action="store",
        type=str,
        default=None,
        help='Elastic index'
    )

    return options


def Main():
    ###GET OPTIONS###
    arguements = GetOptions()
    options = arguements.parse_args()
    
    if options.listtz_flag is True:
        for tz in TIMEZONES:
            print tz
            
        sys.exit(0)
        
    CheckOutputOptions(options)
    CheckImputOptions(options)
    
    if options.jmp_flag == True:
        jmpHandler = JmpHandler(options)
        jmpHandler.ParseJmpFiles()
    else:
        lnkHandler = LnkHandler(options)
        lnkHandler.ParseLinkFiles()
    
def CheckImputOptions(options):
    if options.file_name is None:
        if options.pipe_flag == False:
            print 'No source. Use -f OR --pipe'
            sys.exit(1)
    
def CheckOutputOptions(options):
    if options.csv_flag == False:
        if options.json_flag == False:
            if options.eshost is None:
                print 'No output options used. Use --json OR --txt OR --eshost "ESHOSTIP" --index INDEX'
                sys.exit(1)
 
class JmpHandler():
    def __init__(self,options):
        self.options = options
        self.files_to_parse = []
        
        self.AppIds = self.LoadAppList()
        
        if options.pipe_flag:
            for line in sys.stdin:
                line = line.strip('\n')
                self.files_to_parse.append(line)
        else:
            self.files_to_parse.append(options.file_name)
            
        logging.info('Initialized Jumplist Handler')
        
    def LoadAppList(self):
        applist = {}
        filehandle = None
        if os.path.isfile('AppIdList.txt'):
            filehandle = open('AppIdList.txt','r')
        else:
            filehandle = StringIO.StringIO(APPID_STR)
            
        for line in filehandle:
            line = line.strip("\n")
            fields = line.split("\t")
            
            applist[fields[0]] = fields[1]
            
        return applist
        
    def ParseJmpFiles(self):
        
        outHandler = OutputHandler(
            self.options
        )
        
        outHandler.WriteHeader()
        
        if self.options.eshost is not None:
            outHandler.CreateElasticIndex(self.options)
        
        for jmpfilename in self.files_to_parse:
            logging.info('Parsing File: {}'.format(jmpfilename))
            
            try:
                thisJmp = JmpFile(
                    jmpfilename,
                    self.options,
                    outHandler,
                    app_ids = self.AppIds
                )
            except IOError:
                logging.error('IOError on file {}'.format(jmpfilename))
                continue
            
        if self.options.eshost is not None:
            outHandler.ElasticBulkInsert(
                outHandler.es_records
            )
            
        outHandler.WriteFooter()
    
class LnkHandler():
    def __init__(self,options,filehandle=None,filename=None,jmp_info={'AppIdCode':None,'AppIdName':None}):
        self.options = options
        self.filename = filename
        self.filehandle = filehandle
        self.jmp_info = jmp_info
        
        self.files_to_parse = []
        
        if self.filehandle is None:
            if options.pipe_flag:
                for line in sys.stdin:
                    line = line.strip('\n')
                    self.files_to_parse.append(line)
            else:
                self.files_to_parse.append(options.file_name)
        else:
            self.files_to_parse.append(self.filehandle)
            
        logging.info('Initialized Link Handler')
    
    def ParseLinkFiles(self):
        
        outHandler = OutputHandler(
            self.options
        )
        
        outHandler.WriteHeader()
        
        if self.options.eshost is not None:
            outHandler.CreateElasticIndex(self.options)
        
        for lnkfilename in self.files_to_parse:
            self.ParseLinkFile(
                lnkfilename,
                outHandler
            )
            
        if self.options.eshost is not None:
            outHandler.ElasticBulkInsert(
                outHandler.es_records
            )
            
        outHandler.WriteFooter()
        
    def ParseLinkFile(self,lnkfilename,outHandler):
        if isinstance(lnkfilename,StringIO.StringIO):
            logging.info('Parsing Link in File: {}'.format(self.filename))
        else:
            logging.info('Parsing File: {}'.format(lnkfilename))
            
        if self.filehandle is None:
            try:
                thisLf = LnkFile(
                    filename=lnkfilename,
                    options=self.options
                )
            except IOError:
                logging.error('IOError on file {}'.format(lnkfilename))
                return 0
        else:
            lnkfilename = self.filename
            
            try:
                thisLf = LnkFile(
                    filehandle=self.filehandle,
                    filename=self.filename,
                    options=self.options
                )
            except IOError:
                logging.error('IOError on file {}'.format(lnkfilename))
                return 0
        
        record = thisLf._GetRecordInfo()
        
        record.update(
            self.jmp_info
        )
        
        outHandler.WriteRecord(record)
        
        if self.options.eshost is not None:
            record['EntryNames'] = thisLf.EntryNames
            record['EntryReferences'] = thisLf.EntryReferences
            record['ExternalTag'] = thisLf.ExternalTag
            
            #Add Timestamp#
            timestamp = datetime.datetime.now()
            
            record.update({
                'index_timestamp': timestamp.strftime("%m/%d/%Y %H:%M:%S.%f")
            })
            
            #Create hash of our record to be the id#
            m = md5.new()
            m.update(json.dumps(record))
            hash_id = m.hexdigest()
            
            action = {
                "_index": self.options.index,
                "_type": 'gc_link_file',
                "_id": hash_id,
                "_source": record
            }
            
            outHandler.AppendAction(
                action
            )
    

class OutputHandler():
    def __init__(self,options):
        self.options = options
        if self.options.csv_flag == True:
            self.writer = csv.DictWriter(
                sys.stdout,
                fieldnames=COLUMNS,
                delimiter=self.options.delimiter,
                lineterminator='\n'
            )
        
        if self.options.json_flag == True:
            pass
        
        pass
    
    def AppendAction(self,action):
        self.es_records.append(
            action
        )
    
    def CreateElasticIndex(self,options):
        self.es_records = []
        
        self.esConfig = elastichandler.EsConfig(
            host=options.eshost
        )
        
        self.esHandler = elastichandler.EsHandler(
            self.esConfig
        )
        
        if options.index is not None:
            result = self.esHandler.CheckForIndex(
                options.index
            )
            
            if result == False:
                self.esHandler.InitializeIndex(
                    options.index
                )
        else:
            print 'No Index specified'
            sys.exit()
            
        #Check if mapping exists#
        result = self.esHandler.CheckForMapping(
            'gc_link_file',
            index=options.index
        )
        
        #Create mapping if not exists#
        if result == False:
            self.esHandler.InitializeMapping(
                'gc_link_file',
                LINKMAPPING,
                index=options.index
            )
    
    def ElasticBulkInsert(self,records):
        self.esHandler.BulkIndexRecords(
            records
        )
    
    def WriteHeader(self):
        if self.options.csv_flag == True:
            self.writer.writeheader()
            
        if self.options.json_flag == True:
            self.records = []
    
    def WriteRecord(self,record):
        if self.options.csv_flag == True:
            thisRow = record
            
            if 'DestInfo' in record:
                thisRow.update(record['DestInfo'])
            if record['LnkTrgData'] is not None:
                thisRow.update(record['LnkTrgData'])
            
            out_row = {}
            for column in COLUMNS:
                if column in record:
                    out_row[column] = record[column]
                else:
                    out_row[column] = None
            
            self.writer.writerow(out_row)
            
        if self.options.json_flag == True:
            self.records.append(record)
    
    def WriteFooter(self):
        if self.options.csv_flag == True:
            pass
            
        if self.options.json_flag == True:
            json_out = {'gc_link_file':self.records}
            print json.dumps(json_out,indent=4)
    
class DestList(dict):
    def __init__(self,options,buf):
        self['header'] = DestListHeader(
            buf[:32]
        )
        self['entries'] = []
        start_ofs = 32
        for entry_num in range(1,self['header']['entry_count']+1):
            destList = DestListEntry(
                options,
                buf[start_ofs:]
            )
            entry_size = 114 + destList['DLE_PathSize']
            start_ofs += entry_size
            self['entries'].append(destList)
        
class DestListHeader(dict):
    def __init__(self,buf):
        if buf is not None:
            self['version'] = struct.unpack("<L", buf[0:4])[0]
            self['entry_count'] = struct.unpack("<L", buf[4:8])[0]
            self['pinned_entry_count'] = struct.unpack("<L", buf[8:12])[0]
            self['unknown_1'] = struct.unpack("<L", buf[12:16])[0]
            self['unknown_2'] = struct.unpack("<L", buf[16:20])[0]
            self['unknown_3'] = struct.unpack("<L", buf[20:24])[0]
            self['unknown_4'] = struct.unpack("<L", buf[24:28])[0]
            self['unknown_5'] = struct.unpack("<L", buf[28:32])[0]

class DestListEntry(dict):
    def __init__(self,options,buf):
        if buf is not None:
            self['DLE_Unknown1Hash'] = buf[0:8].encode('hex')
            self['DLE_DroidVolId'] = buf[8:8+16].encode('hex')
            self['DLE_DroidFileId'] = buf[24:24+16].encode('hex')
            self['DLE_BirthDroidVolId'] = buf[40:40+16].encode('hex')
            self['DLE_BirthDroidFileId'] = buf[56:56+16].encode('hex')
            self['DLE_Hostname'] = buf[72:72+16]
            self['DLE_EntryNum'] = struct.unpack("<L", buf[88:88+4])[0]
            self['DLE_Unknown2'] = struct.unpack("<L", buf[92:92+4])[0]
            self['DLE_Unknown3'] = struct.unpack("<L", buf[96:96+4])[0]
            self['DLE_LastModificationDatetime'] = ConvertDateTime(
                options.timeformat,
                options.timezone,
                GetTimeStamp(buf[100:100+8])
            )
            self['DLE_PinStatus'] = struct.unpack("<L", buf[108:108+4])[0]
            self['DLE_PathSize'] = struct.unpack("<H", buf[112:112+2])[0] * 2
            self['DLE_Path'] = buf[114:114+self['DLE_PathSize']].decode('utf-16le')


class JmpFile():
    def __init__(self,jmpfilename,options,outHandler,app_ids=None):
        logging.info('Linkfile: {}'.format(options.file_name))
        self.outHandler = outHandler
        self.options = options
        self.timezone = options.timezone
        self.file_name = jmpfilename
        self.timeformat = options.timeformat
        self.error_flag = False
        self.error_message = ''
        self.ExternalTag = None
        
        self.appinfo = self.GetAppId(self.file_name,app_ids)
        
        self.type = None
        
        if pyolecf.check_file_signature(self.file_name):
            self.type = 'AutomaticDestinations'
            self.jmpFile = pyolecf.file()
            self.jmpFile.open(self.file_name)
            self._GetItems()
        else:
            self.type = 'CustomDestinations'
            self.jmpFile = open(self.file_name,'rb')
        
            self._FindLnkFiles()
    
    def _FindLnkFiles(self):
        raw_file_data = self.jmpFile.read()
        item_cnt = 0
        for match in re.finditer('\x4c\x00\x00\x00', raw_file_data):
            start_offset = match.regs[0][0]
            
            dfh = StringIO.StringIO(raw_file_data[start_offset:])
            
            lnk_check = pylnk.check_file_signature_file_object(
                dfh
            )
            
            if lnk_check:
                lnkHandler = LnkHandler(
                    self.options,
                    filename=self.file_name,
                    filehandle=dfh,
                    jmp_info = self.appinfo
                )
                
                lnkHandler.ParseLinkFile(
                    dfh,
                    self.outHandler
                )
            else:
                logging.error('item {} is not lnk struct at offset {} [Name:{}]'.format(
                    item_cnt,
                    start_offset,
                    item.name
                ))
                item_cnt += 1
                continue
            
            item_cnt += 1
    
    def _GetItems(self):
        '''
        Issues here. Skiping first link entry
        '''
        item_cnt = 0
        
        dest_list = self.jmpFile.root_item.get_sub_item_by_name('DestList')
        dsize = dest_list.get_size()
        data = dest_list.read(dsize)
        data_hex = data.encode('hex')
        
        dest_list = DestList(
            self.options,
            data
        )

        for entry in dest_list['entries']:
            item_name = hex(entry['DLE_EntryNum'])[2:]
            item = self.jmpFile.root_item.get_sub_item_by_name(
                item_name
            )
            self.appinfo['DestInfo'] = entry
            dsize = item.get_size()
            data = item.read(dsize)
            data_hex = data.encode('hex')
            
            dfh = StringIO.StringIO(data)
            
            lnk_check = pylnk.check_file_signature_file_object(
                dfh
            )
            
            if lnk_check:
                lnkHandler = LnkHandler(
                    self.options,
                    filename=self.file_name,
                    filehandle=dfh,
                    jmp_info = self.appinfo
                )
                
                lnkHandler.ParseLinkFile(
                    dfh,
                    self.outHandler
                )
            else:
                logging.error('item {} is not lnk struct [Name:{}]'.format(
                    item_cnt,
                    item.name
                ))
                item_cnt += 1
                continue
            
            item_cnt += 1
            
    def GetAppId(self,jmpfilename,AppIds):
        info = {
            'AppIdCode':None,
            'AppIdName':None
        }
        pattern = re.compile("([0-9a-f]{16})\.(custom|automatic)Destinations\-ms$") #5d696d521de238c3.customDestinations-ms
        match = pattern.search(jmpfilename)
        
        if match is not None:
            appid = match.group(1)
            info['AppIdCode'] = appid
            if appid in AppIds:
                info['AppIdName'] = AppIds[appid]
            thistype = match.group(2)
        
        return info
    
class LnkFile():
    def __init__(self,filename=None,filehandle=None,options=None):
        logging.info('Linkfile: {}'.format(options.file_name))
        self.options = options
        self.timezone = options.timezone
        self.file_name = filename
        
        if filehandle is None:
            self.lnkFile = pylnk.file()
            self.lnkFile.open(
                self.file_name
            )
        else:
            self.lnkFile = pylnk.file()
        
            self.lnkFile.open_file_object(
                filehandle
            )
        
        self.timeformat = options.timeformat
        self.error_flag = False
        self.error_message = ''
        
        self.EntryNames = []
        self.EntryReferences = []
        self.ExternalTag = None
        
        if not self.error_flag:
            self._GetShellItems()
        
    def _GetShellItems(self):
        self.shellItems = pyfwsi.item_list()
        
        if self.lnkFile.link_target_identifier_data is not None:
            link_target_identifier_data_hex = self.lnkFile.link_target_identifier_data.encode('hex')
        
        if self.lnkFile.link_target_identifier_data is None:
            self.shellItems = None
        else:
            try:
                self.shellItems.copy_from_byte_stream(
                    self.lnkFile.link_target_identifier_data
                )
            except Exception as e:
                logging.error('Shell items not parsed. {}'.format(e.message))
                self.shellItems = self.lnkFile.link_target_identifier_data.encode('hex')
                
        return 1
    
    def _GetRecordInfo(self):
        record = {
            'Source':self.file_name,
            'Codepage':self.lnkFile.ascii_codepage,
            'CmdArgs':self.lnkFile.command_line_arguments,
            'Description':self.lnkFile.description,
            'DriveSerialNumber':self._FormatVolumeSerialNum(self.lnkFile.drive_serial_number),
            'DriveType':self.EnumDriveType(
                self.lnkFile.drive_type
            ),
            'EnvVarLoc':self.lnkFile.environment_variables_location,
            'AccessDateTime':ConvertDateTime(
                self.timeformat,
                self.timezone,
                self.lnkFile.file_access_time
            ),
            'Flags':EnumerateFlags(
                self.lnkFile.file_attribute_flags,
                FILE_ATTRIBUTE_FLAGS
            ),
            'CreationDateTime':ConvertDateTime(
                self.timeformat,
                self.timezone,
                self.lnkFile.file_creation_time
            ),
            'ModificationDateTime':ConvertDateTime(
                self.timeformat,
                self.timezone,
                self.lnkFile.file_modification_time
            ),
            'FileSize':self.lnkFile.file_size,
            'IconLoc':self.lnkFile.icon_location,
            'LocalPath':self.lnkFile.local_path,
            'NetworkPath':self.lnkFile.network_path,
            'RelativePath':self.lnkFile.relative_path,
            'VolumeLabel':self.lnkFile.volume_label,
            'WorkingDir':self.lnkFile.working_directory
        }
        
        #Get Base Name and Ext#
        record['BaseName'] = self._GetFileName([
            record['LocalPath'],
            record['NetworkPath'],
            record['RelativePath']
        ])
        
        record['FileExt'] = None
        if record['BaseName'] is not None:
            record['FileExt'] = (os.path.splitext(record['BaseName']))[1]
        if record['FileExt'] is not None:
            record['FileExt'] = record['FileExt'].strip('.')
        
        items = self._GetTrgInfo()
        
        record['LnkTrgData'] = items
        
        ###Create Tags###
        if record['LnkTrgData'] is None:
            if record['NetworkPath'] is not None:
                fullpath = record['NetworkPath'].strip('\\')
                entries = fullpath.split('\\')
                self.EntryNames = entries
                self.ExternalTag = 'Network'
        
        return record
    
    def _GetFileName(self,filename_list):
        filename = None
        
        for name in filename_list:
            if name is not None:
                if name != '':
                    filename = os.path.basename(name)
                    break
        
        return filename
    
    def _FormatVolumeSerialNum(self,number):
        vol_s = None
        
        if number is not None:
            vol_s = hex(number)[2:-1].zfill(8)
            vol_s = vol_s[:4] + '-' + vol_s[4:]
            
        return vol_s
    
    def _GetTrgInfo(self):
        if self.shellItems is None:
            return None
        
        if type(self.shellItems) == str:
            return self.shellItems
        
        aliasnames = {
            'root_folder':'RootFolder',
            'volume':'Volume',
            'file_entry':'FileEntries',
            'item':'FileEntries',
            'network_location':'NetworkLocation'
        }
        record = {}
        summary = {}
        
        si_count = 1
        dist_class_types_str = {}
        dist_class_types_hex = {}
        list_ext_block_types = []
        for shell_item in self.shellItems.items:
            #Create Info Container for Clean Output#
            shell_info = {}
            
            #Set Types#
            shell_info['ShellItemTypeStr'] = type(shell_item).__name__
            shell_info['ShellItemTypeHex'] = hex(shell_item.class_type)
            
            #Set Distinct Shell Item Class Type#
            dist_class_types_str[shell_info['ShellItemTypeStr']] = True
            dist_class_types_hex[shell_info['ShellItemTypeHex']] = True
            
            ###Get Extention Block Info###
            shell_info['ExtentionBlocks'] = []
            shell_info['ExtentionBlockCount'] = shell_item.number_of_extension_blocks
            if shell_item.number_of_extension_blocks > 0:
                if shell_item.number_of_extension_blocks > 1:
                    logging.info('Multiple Extention Blocks in {}'.format(self.file_name))
                    
                for extention_blk in shell_item.extension_blocks:
                    list_ext_block_types.append(
                        hex(extention_blk.signature)
                    )
                    ext_blk_info = self._GetExtentionAttribs(
                        si_count-1,
                        extention_blk
                    )
                    
                    shell_info['ExtentionBlocks'].append(
                        ext_blk_info
                    )
            
            if (shell_info['ShellItemTypeStr'] == 'volume' or
              shell_info['ShellItemTypeStr'] == 'file_entry' or
              shell_info['ShellItemTypeStr'] == 'item' or
              shell_info['ShellItemTypeStr'] == 'root_folder' or
              shell_info['ShellItemTypeStr'] == 'network_location' or
              shell_info['ShellItemTypeStr'] == 'identifier'):
                shell_info['ShellFolderId'] = getattr(shell_item,'shell_folder_identifier', None)
                shell_info['Name'] = getattr(shell_item,'name', None)
                shell_info['FileSize'] = getattr(shell_item,'file_size', None)
                shell_info['Identifier'] = getattr(shell_item,'identifier',None)
                shell_info['Location'] = getattr(shell_item,'location',None)
                shell_info['Description'] = getattr(shell_item,'description',None)
                shell_info['Comments'] = getattr(shell_item,'comments',None)
                
                try:
                    shell_info['ModificationTime'] = ConvertDateTime(
                        self.timeformat,
                        self.timezone,
                        getattr(shell_item,'modification_time', None)
                    )
                except IOError as e:
                    logging.warn('{}'.format(e.message))
                    shell_info['ModificationTime'] = self._GetNullDateTime()
                
                if aliasnames[shell_info['ShellItemTypeStr']] not in record:
                    record[aliasnames[shell_info['ShellItemTypeStr']]] = []
                    
                record[aliasnames[shell_info['ShellItemTypeStr']]].append(shell_info)
            else:
                logging.warn('unhandled shell item type: {}'.format(shell_info['ShellItemTypeStr']))
                raise Exception('unhandled shell item type: {}'.format(shell_info['ShellItemTypeStr']))
                
            si_count += 1
        
        summary['ItemCount'] = si_count
        summary['DistinctTypesStr'] = ';'.join(dist_class_types_str.keys())
        summary['DistinctTypesHex'] = ';'.join(dist_class_types_hex.keys())
        summary['ExtentionListing'] = ';'.join(list_ext_block_types)
        list_ext_block_types
        
        record.update(summary)
        
        ###Get Parent Info###
        parentinfo = {
            'ParentLongName':None,
            'ParentEntryNum':None,
            'ParentSeqNum':None
        }
        if 'FileEntries' in record and len(record['FileEntries']) > 1:
            entry_cnt = len(record['FileEntries'])
            try:
                parentinfo['ParentLongName'] = record['FileEntries'][entry_cnt - 2]['ExtentionBlocks'][0]['LongName']
                parentinfo['ParentEntryNum'] = record['FileEntries'][entry_cnt - 2]['ExtentionBlocks'][0]['EntryNum']
                parentinfo['ParentSeqNum'] = record['FileEntries'][entry_cnt - 2]['ExtentionBlocks'][0]['SeqNum']
                parentinfo['ParentRefStr'] = '{}-{}'.format(parentinfo['ParentEntryNum'],parentinfo['ParentSeqNum'])
            except:
                parentinfo['ParentLongName'] = None
                parentinfo['ParentEntryNum'] = None
                parentinfo['ParentSeqNum'] = None
                parentinfo['ParentRefStr'] = None
            
        record.update(parentinfo)
        
        return record
    
    def _GetExtentionAttribs(self,shell_item_pos,extention_block):
        info = {}
        
        sig = hex(extention_block.signature)
        
        #We can do extention block stats here for research#
        # TODO #
        
        #if extention_block.signature != 3203334148:
        #    logging.error(
        #        "Extention Block {} thats not 4".format(
        #            hex(extention_block.signature)
        #        )
        #    )
        
        if hasattr(extention_block,'access_time'):
            try:
                info['AccessTime'] = ConvertDateTime(
                    self.timeformat,
                    self.timezone,
                    extention_block.access_time
                )
            except IOError as e:
                logging.warn('{}'.format(e.message))
                info['AccessTime'] = self._GetNullDateTime()
        
        if hasattr(extention_block,'creation_time'):
            try:
                info['CreationTime'] = ConvertDateTime(
                    self.timeformat,
                    self.timezone,
                    extention_block.creation_time
                )
            except IOError as e:
                logging.warn('{}'.format(e.message))
                info['CreationTime'] = self._GetNullDateTime()
        
        if hasattr(extention_block,'data_size'): info['ExtentionBlockSize'] = extention_block.data_size
        if hasattr(extention_block,'file_reference'):
            info['FileReferenceInt'] = extention_block.file_reference
            info.update(
                self._GetRefNums(
                    extention_block.file_reference
                )
            )
            
        if hasattr(extention_block,'localized_name'): info['LocalizedName'] = extention_block.localized_name
        
        if hasattr(extention_block,'long_name'):
            info['LongName'] = extention_block.long_name
            if info['LongName'] != '':
                self.EntryNames.append(info['LongName'])
            
        info['Signature'] = hex(extention_block.signature)
        
        return info
    
    def _GetNullDateTime(self):
        datetime_in = datetime.datetime(1601,1,1)
        
        try:
            datetime_out_str = datetime_in.strftime(
                self.timeformat
            )
        except Exception as e:
            datetime_out_str = datetime_in.isoformat(" ").split("+")[0]
            #datetime_out_str = 'na'
        
        return datetime_out_str
    
    def _GetRefNums(self,lref):
        info = {}
        
        if lref is not None:
            raw = struct.pack('Q',lref)
            
            info['EntryNum'] = struct.unpack("<Lxx", raw[:6])[0]
            info['SeqNum'] = struct.unpack("<H", raw[6:8])[0]
            info['RefNum'] = '{}-{}'.format(info['EntryNum'],info['SeqNum'])
            self.EntryReferences.append(info['RefNum'])
        else:
            info['EntryNum'] = None
            info['SeqNum'] = None
            info['RefNum'] = None
            
        return info
    
    def EnumDriveType(self,flag):
        if flag is not None:
            return DRIVE_TYPES[flag]
        
        return None
    
def EnumerateFlags(flag,flag_mapping):
    str_flag = ''
        
    for i in flag_mapping:
        if (i & flag):
            str_flag = ''.join([str_flag,flag_mapping[i]])

    return str_flag

def GetTimeStamp(raw_timestamp):
    timestamp = struct.unpack("Q",raw_timestamp)[0]
    
    if datetime < 0:
        return None
    
    microsecs, _ = divmod(timestamp, 10)
    timeDelta = datetime.timedelta(
        microseconds=microsecs
    )
    
    origDateTime = datetime.datetime(1601, 1, 1)
    
    new_datetime = origDateTime + timeDelta
  
    return new_datetime

def ConvertDateTime(timeformat,timezone,datetime_in):
    if datetime_in is None:
        return None
    
    utc = pytz.timezone("UTC")
    new_tz = pytz.timezone(timezone)
    
    datetime_out = new_tz.localize(datetime_in).astimezone(utc)
    
    datetime_out_str = timeformat.format(datetime_out)
    
    return datetime_out_str

if __name__ == '__main__':
    Main()