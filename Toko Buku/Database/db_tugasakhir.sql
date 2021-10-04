-- phpMyAdmin SQL Dump
-- version 5.0.2
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Waktu pembuatan: 29 Jun 2021 pada 08.39
-- Versi server: 10.4.13-MariaDB
-- Versi PHP: 7.4.7

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `db_tugasakhir`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `tb_kasir`
--

CREATE TABLE `tb_kasir` (
  `no` int(5) NOT NULL,
  `tglPembelian` varchar(20) DEFAULT NULL,
  `jenisBuku` varchar(20) NOT NULL,
  `namaBuku` varchar(50) NOT NULL,
  `hargaBuku` varchar(10) NOT NULL,
  `jumlahBuku` varchar(4) NOT NULL,
  `jenisDiskon` varchar(5) NOT NULL,
  `totalHarga` varchar(10) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Dumping data untuk tabel `tb_kasir`
--

INSERT INTO `tb_kasir` (`no`, `tglPembelian`, `jenisBuku`, `namaBuku`, `hargaBuku`, `jumlahBuku`, `jenisDiskon`, `totalHarga`) VALUES
(1, '29/06/2021', 'IT', 'PHP Pemula', '78000', '1', '10%', '70200'),
(2, '29/06/2021', 'Agama', 'Kisah 25 Nabi', '50000', '2', '25%', '75000');

-- --------------------------------------------------------

--
-- Struktur dari tabel `tb_stokbarang`
--

CREATE TABLE `tb_stokbarang` (
  `no` int(5) NOT NULL,
  `tgl` varchar(20) NOT NULL,
  `jenisBuku` varchar(20) NOT NULL,
  `namaBuku` varchar(50) NOT NULL,
  `hargaBuku` varchar(10) NOT NULL,
  `jumlahBuku` varchar(4) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Dumping data untuk tabel `tb_stokbarang`
--

INSERT INTO `tb_stokbarang` (`no`, `tgl`, `jenisBuku`, `namaBuku`, `hargaBuku`, `jumlahBuku`) VALUES
(1, '29/06/2021', 'IT', 'PHP Pemula', '78000', '75'),
(2, '29/06/2021', 'Agama', 'Kisah 25 Nabi', '50000', '100');

--
-- Indexes for dumped tables
--

--
-- Indeks untuk tabel `tb_kasir`
--
ALTER TABLE `tb_kasir`
  ADD PRIMARY KEY (`no`);

--
-- Indeks untuk tabel `tb_stokbarang`
--
ALTER TABLE `tb_stokbarang`
  ADD PRIMARY KEY (`no`);

--
-- AUTO_INCREMENT untuk tabel yang dibuang
--

--
-- AUTO_INCREMENT untuk tabel `tb_kasir`
--
ALTER TABLE `tb_kasir`
  MODIFY `no` int(5) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;

--
-- AUTO_INCREMENT untuk tabel `tb_stokbarang`
--
ALTER TABLE `tb_stokbarang`
  MODIFY `no` int(5) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
