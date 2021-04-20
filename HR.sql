-- MySQL dump 10.13  Distrib 8.0.23, for Win64 (x86_64)
--
-- Host: localhost    Database: otd_kadrov
-- ------------------------------------------------------
-- Server version	8.0.23

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!50503 SET NAMES utf8mb4 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `akadem_otpusk`
--

DROP TABLE IF EXISTS `akadem_otpusk`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `akadem_otpusk` (
  `id` int NOT NULL AUTO_INCREMENT,
  `value` varchar(5) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `id_UNIQUE` (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `akadem_otpusk`
--

LOCK TABLES `akadem_otpusk` WRITE;
/*!40000 ALTER TABLE `akadem_otpusk` DISABLE KEYS */;
/*!40000 ALTER TABLE `akadem_otpusk` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `city`
--

DROP TABLE IF EXISTS `city`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `city` (
  `id` int NOT NULL AUTO_INCREMENT,
  `city` varchar(100) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `id_UNIQUE` (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `city`
--

LOCK TABLES `city` WRITE;
/*!40000 ALTER TABLE `city` DISABLE KEYS */;
/*!40000 ALTER TABLE `city` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `comm`
--

DROP TABLE IF EXISTS `comm`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `comm` (
  `id` int NOT NULL AUTO_INCREMENT,
  `type` varchar(20) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `id_UNIQUE` (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `comm`
--

LOCK TABLES `comm` WRITE;
/*!40000 ALTER TABLE `comm` DISABLE KEYS */;
/*!40000 ALTER TABLE `comm` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `district`
--

DROP TABLE IF EXISTS `district`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `district` (
  `id` int NOT NULL AUTO_INCREMENT,
  `district` varchar(100) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `id_UNIQUE` (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `district`
--

LOCK TABLES `district` WRITE;
/*!40000 ALTER TABLE `district` DISABLE KEYS */;
/*!40000 ALTER TABLE `district` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `goden`
--

DROP TABLE IF EXISTS `goden`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `goden` (
  `id` int NOT NULL AUTO_INCREMENT,
  `Goden` varchar(8) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `id_UNIQUE` (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `goden`
--

LOCK TABLES `goden` WRITE;
/*!40000 ALTER TABLE `goden` DISABLE KEYS */;
/*!40000 ALTER TABLE `goden` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `group`
--

DROP TABLE IF EXISTS `group`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `group` (
  `id` int NOT NULL AUTO_INCREMENT,
  `groupName` varchar(11) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `id_UNIQUE` (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `group`
--

LOCK TABLES `group` WRITE;
/*!40000 ALTER TABLE `group` DISABLE KEYS */;
/*!40000 ALTER TABLE `group` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `otdelenie`
--

DROP TABLE IF EXISTS `otdelenie`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `otdelenie` (
  `id` int NOT NULL AUTO_INCREMENT,
  `number` varchar(2) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `id_UNIQUE` (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `otdelenie`
--

LOCK TABLES `otdelenie` WRITE;
/*!40000 ALTER TABLE `otdelenie` DISABLE KEYS */;
/*!40000 ALTER TABLE `otdelenie` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `sex`
--

DROP TABLE IF EXISTS `sex`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `sex` (
  `id` int NOT NULL AUTO_INCREMENT,
  `sex` varchar(10) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `id_UNIQUE` (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `sex`
--

LOCK TABLES `sex` WRITE;
/*!40000 ALTER TABLE `sex` DISABLE KEYS */;
INSERT INTO `sex` VALUES (1,'Муж.'),(2,'Жен.'),(3,' ');
/*!40000 ALTER TABLE `sex` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `student`
--

DROP TABLE IF EXISTS `student`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `student` (
  `id` int NOT NULL AUTO_INCREMENT,
  `name` varchar(51) COLLATE utf8_unicode_ci DEFAULT NULL,
  `birth` date DEFAULT NULL,
  `id_sex` int DEFAULT NULL,
  `indx` int DEFAULT NULL,
  `id_city` int DEFAULT NULL,
  `country` varchar(51) COLLATE utf8_unicode_ci DEFAULT NULL,
  `id_district` int DEFAULT NULL,
  `street` varchar(51) COLLATE utf8_unicode_ci DEFAULT NULL,
  `house` varchar(11) COLLATE utf8_unicode_ci DEFAULT NULL,
  `flat` int DEFAULT NULL,
  `phone` varchar(21) COLLATE utf8_unicode_ci DEFAULT NULL,
  `passpSeries` varchar(21) COLLATE utf8_unicode_ci DEFAULT NULL,
  `passpNumber` varchar(21) COLLATE utf8_unicode_ci DEFAULT NULL,
  `passpKemVidan` varchar(201) COLLATE utf8_unicode_ci DEFAULT NULL,
  `passpDate` date DEFAULT NULL,
  `idGroup` int DEFAULT NULL,
  `id_otdelenie` int DEFAULT NULL,
  `id_comm` int DEFAULT NULL,
  `prikazNumIn` varchar(21) COLLATE utf8_unicode_ci DEFAULT NULL,
  `dateIn` date DEFAULT NULL,
  `prikazNumOut` varchar(21) COLLATE utf8_unicode_ci DEFAULT NULL,
  `dateOut` date DEFAULT NULL,
  `prichinaOut` varchar(151) COLLATE utf8_unicode_ci DEFAULT NULL,
  `tabNum` int DEFAULT NULL,
  `kval` varchar(51) COLLATE utf8_unicode_ci DEFAULT NULL,
  `prikazNumKval` varchar(51) COLLATE utf8_unicode_ci DEFAULT NULL,
  `kodDoc` int DEFAULT NULL,
  `note` varchar(200) COLLATE utf8_unicode_ci DEFAULT NULL,
  `id_goden` int DEFAULT NULL,
  `katGodnost` varchar(5) COLLATE utf8_unicode_ci DEFAULT NULL,
  `id_academ` int DEFAULT NULL,
  `ciizenship` varchar(30) COLLATE utf8_unicode_ci DEFAULT NULL,
  `homePhone` varchar(21) COLLATE utf8_unicode_ci DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `id_UNIQUE` (`id`),
  KEY `id_sex_idx` (`id_sex`),
  KEY `id_city_idx` (`id_city`),
  KEY `id_district_idx` (`id_district`),
  KEY `id_otdelenie_idx` (`id_otdelenie`),
  KEY `id_group_idx` (`idGroup`),
  KEY `id_comm_idx` (`id_comm`),
  KEY `id_goden_idx` (`id_goden`),
  KEY `id_academ_idx` (`id_academ`),
  CONSTRAINT `id_academ` FOREIGN KEY (`id_academ`) REFERENCES `akadem_otpusk` (`id`),
  CONSTRAINT `id_city` FOREIGN KEY (`id_city`) REFERENCES `city` (`id`),
  CONSTRAINT `id_comm` FOREIGN KEY (`id_comm`) REFERENCES `comm` (`id`),
  CONSTRAINT `id_district` FOREIGN KEY (`id_district`) REFERENCES `district` (`id`),
  CONSTRAINT `id_goden` FOREIGN KEY (`id_goden`) REFERENCES `goden` (`id`),
  CONSTRAINT `id_group` FOREIGN KEY (`idGroup`) REFERENCES `group` (`id`),
  CONSTRAINT `id_otdelenie` FOREIGN KEY (`id_otdelenie`) REFERENCES `otdelenie` (`id`),
  CONSTRAINT `id_sex` FOREIGN KEY (`id_sex`) REFERENCES `sex` (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `student`
--

LOCK TABLES `student` WRITE;
/*!40000 ALTER TABLE `student` DISABLE KEYS */;
/*!40000 ALTER TABLE `student` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2021-04-20 12:17:15
