# Host: localhost  (Version 5.0.45-community-nt)
# Date: 2020-06-30 10:17:43
# Generator: MySQL-Front 5.3  (Build 5.33)

/*!40101 SET NAMES latin1 */;

#
# Structure for table "tb_kriteria"
#

DROP TABLE IF EXISTS `tb_kriteria`;
CREATE TABLE `tb_kriteria` (
  `kd_kriteria` char(4) NOT NULL default '',
  `nm_kriteria` varchar(255) default NULL,
  `jenis_kriteria` varchar(255) default NULL,
  `bobot` double default NULL,
  PRIMARY KEY  (`kd_kriteria`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

#
# Data for table "tb_kriteria"
#

/*!40000 ALTER TABLE `tb_kriteria` DISABLE KEYS */;
INSERT INTO `tb_kriteria` VALUES ('C1','PEKERJAAN','BENEFIT',5),('C2','PENGHASILAN','COST',1),('C3','JUMLAH TANGGUNGAN','BENEFIT',5),('C4','KEPEMILIKAN','BENEFIT',4),('C5','LANTAI','BENEFIT',4),('C6','DINDING','BENEFIT',3);
/*!40000 ALTER TABLE `tb_kriteria` ENABLE KEYS */;

#
# Structure for table "tb_normalisasi"
#

DROP TABLE IF EXISTS `tb_normalisasi`;
CREATE TABLE `tb_normalisasi` (
  `kd_alternatif` char(5) NOT NULL default '0',
  `C1` double default NULL,
  `C2` double default NULL,
  `C3` double default NULL,
  `C4` double default NULL,
  `C5` double default NULL,
  UNIQUE KEY `Id` (`kd_alternatif`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

#
# Data for table "tb_normalisasi"
#


#
# Structure for table "tb_pekerjaan"
#

DROP TABLE IF EXISTS `tb_pekerjaan`;
CREATE TABLE `tb_pekerjaan` (
  `id_pekerjaan` varchar(11) NOT NULL default '',
  `kd_alternatif` char(5) default NULL,
  `pekerjaan` varchar(50) default NULL,
  `penghasilan` varchar(30) default NULL,
  `jm_tanggungan` varchar(2) default NULL,
  `kepemilikan` varchar(30) default NULL,
  `lantai` varchar(30) default NULL,
  `dinding` varchar(25) default NULL,
  PRIMARY KEY  (`id_pekerjaan`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

#
# Data for table "tb_pekerjaan"
#

/*!40000 ALTER TABLE `tb_pekerjaan` DISABLE KEYS */;
INSERT INTO `tb_pekerjaan` VALUES ('0001','A0001','Petani','< 500.000','3','Milik Sendiri','Kayu','Kayu'),('0002','A0002','Wiraswasta','500.000-1.000.000','3','Milik Sendiri','Kayu','Kayu');
/*!40000 ALTER TABLE `tb_pekerjaan` ENABLE KEYS */;

#
# Structure for table "tb_penduduk"
#

DROP TABLE IF EXISTS `tb_penduduk`;
CREATE TABLE `tb_penduduk` (
  `kd_alternatif` char(5) NOT NULL default '',
  `nikk` char(20) default NULL,
  `namakk` varchar(30) default NULL,
  `alamat` varchar(255) default NULL,
  `jk` varchar(25) default NULL,
  `rt` char(3) default NULL,
  `rw` char(3) default NULL,
  PRIMARY KEY  (`kd_alternatif`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

#
# Data for table "tb_penduduk"
#

/*!40000 ALTER TABLE `tb_penduduk` DISABLE KEYS */;
INSERT INTO `tb_penduduk` VALUES ('A0001','033789787326','Dimas','Desa Tellulimpoe,Dusun Lompoe','laki-laki','01','01'),('A0002','033378456235','Toni','Kel.Batu-Batu','laki-laki','02','02');
/*!40000 ALTER TABLE `tb_penduduk` ENABLE KEYS */;

#
# Structure for table "tb_penilaiaan_alternatif"
#

DROP TABLE IF EXISTS `tb_penilaiaan_alternatif`;
CREATE TABLE `tb_penilaiaan_alternatif` (
  `kd_alternatif` char(5) NOT NULL default '0',
  `C1` double default NULL,
  `C2` double default NULL,
  `C3` double default NULL,
  `C4` double default NULL,
  `C5` double default NULL,
  UNIQUE KEY `Id` (`kd_alternatif`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

#
# Data for table "tb_penilaiaan_alternatif"
#


#
# Structure for table "tb_penilaian"
#

DROP TABLE IF EXISTS `tb_penilaian`;
CREATE TABLE `tb_penilaian` (
  `kd_alternatif` char(5) NOT NULL default '0',
  `kd_kriteria` char(4) default NULL,
  `nilai` varchar(255) default NULL,
  UNIQUE KEY `Id` (`kd_alternatif`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

#
# Data for table "tb_penilaian"
#


#
# Structure for table "tb_rangekriteria"
#

DROP TABLE IF EXISTS `tb_rangekriteria`;
CREATE TABLE `tb_rangekriteria` (
  `kd_range` char(5) NOT NULL default '',
  `kd_kriteria` char(5) default NULL,
  `pekerjaan` varchar(25) default NULL,
  `range_kriteria` varchar(255) default NULL,
  `nilairange` double default NULL,
  PRIMARY KEY  (`kd_range`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

#
# Data for table "tb_rangekriteria"
#

/*!40000 ALTER TABLE `tb_rangekriteria` DISABLE KEYS */;
INSERT INTO `tb_rangekriteria` VALUES ('R0001','C1','PEKERJAAN','PNS',1),('R0002','C1','PEKERJAAN','WIRASWASTA',2),('R0003','C1','PEKERJAAN','PETANI',3),('R0004','C1','PEKERJAAN','BURUH',4),('R0005','C1','PEKERJAAN','PENGANGGURAN',5);
/*!40000 ALTER TABLE `tb_rangekriteria` ENABLE KEYS */;

#
# Structure for table "tb_simpanhasil"
#

DROP TABLE IF EXISTS `tb_simpanhasil`;
CREATE TABLE `tb_simpanhasil` (
  `kd_alternatif` char(5) NOT NULL default '0',
  `totalnilai` double default NULL,
  UNIQUE KEY `Id` (`kd_alternatif`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

#
# Data for table "tb_simpanhasil"
#

