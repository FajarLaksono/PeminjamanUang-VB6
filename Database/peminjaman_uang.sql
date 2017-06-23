-- phpMyAdmin SQL Dump
-- version 4.5.1
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: 12 Jun 2017 pada 09.24
-- Versi Server: 10.1.10-MariaDB
-- PHP Version: 7.0.4

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `peminjaman_uang`
--

CREATE DATABASE peminjaman_uang;
USE peminjaman_uang;

-- --------------------------------------------------------

--
-- Struktur dari tabel `tblanggota`
--

CREATE TABLE `tblanggota` (
  `no_ktp` varchar(255) NOT NULL COMMENT 'Nomer KTP anggota',
  `nama` varchar(255) NOT NULL COMMENT 'Nama anggota',
  `tempat_lahir` varchar(255) NOT NULL,
  `tanggal_lahir` date NOT NULL,
  `jenis_kelamin` varchar(10) NOT NULL COMMENT 'Jenis Kelamin Anggota (Laki-Laki/Perempuan)',
  `pekerjaan` varchar(100) NOT NULL COMMENT 'Pekerjaan Anggota',
  `telepon` varchar(20) NOT NULL COMMENT 'Nomer Telepon Anggota',
  `alamat` varchar(255) NOT NULL COMMENT 'Alamat Anggota',
  `kode_pos` int(10) NOT NULL,
  `foto` varchar(255) NOT NULL COMMENT 'Nama Foto Anggota pada direktori'
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `tblanggota`
--

INSERT INTO `tblanggota` (`no_ktp`, `nama`, `tempat_lahir`, `tanggal_lahir`, `jenis_kelamin`, `pekerjaan`, `telepon`, `alamat`, `kode_pos`, `foto`) VALUES
('12151247', 'Fajar Aziz Laksono', 'Purwokerto', '1997-10-19', 'Laki-Laki', 'CEO', '855555555', 'Purwokerto', 53123, 'FotoAnggota29-April-2017-19-59-35FajarAzizLaksono.jpg'),
('1267421', 'Bruce Wayne', 'Gotham', '2001-07-20', 'Laki-Laki', 'Batman', '085813883523', 'Gotham', 9090999, ''),
('99988237109843', 'Rosalio Cahyo Romadhon', 'purwokerto', '1999-07-08', 'Laki-Laki', 'Wonder Woman', '85', 'mau tau aja 1 / 2 konoha asd asd', 53121, 'FotoAnggota23-May-2017-15-10-197dfd5bfa58d09c66a6a49251b882bcdd.jpg');

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbljaminan`
--

CREATE TABLE `tbljaminan` (
  `id_jaminan` int(11) NOT NULL,
  `jenis` varchar(25) NOT NULL,
  `foto` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `tbljaminan`
--

INSERT INTO `tbljaminan` (`id_jaminan`, `jenis`, `foto`) VALUES
(5, '98923', 'jaminan11-Mei-2017-20-43-36140910_BURKARD_82703_head.jpg'),
(6, '123', 'jaminan12-Mei-2017-15-31-26140910_BURKARD_82703_head.jpg'),
(7, 'Pop', 'jaminan12-Mei-2017-15-33-25140910_BURKARD_82703_head.jpg'),
(8, '23123', 'jaminan12-Mei-2017-15-34-89c9337d81e9e3606de35e2924672d9cb.jpg'),
(9, '123', 'jaminan12-Mei-2017-15-36-27140910_BURKARD_82703_head.jpg'),
(10, '1', 'jaminan12-Mei-2017-15-38-15140910_BURKARD_82703_head.jpg'),
(11, 'Bat Cave', 'jaminan26-Mei-2017-18-34-117dfd5bfa58d09c66a6a49251b882bcdd.jpg'),
(12, 'Bat Cave', 'jaminan26-Mei-2017-18-38-137dfd5bfa58d09c66a6a49251b882bcdd.jpg'),
(13, '123', 'jaminan28-Mei-2017-18-56-587dfd5bfa58d09c66a6a49251b882bcdd.jpg'),
(14, '123', 'jaminan28-Mei-2017-19-56-55wallpapers-sci-fi-0.jpg'),
(15, 'iuuiu', 'jaminan28-Mei-2017-19-57-17121806.jpg'),
(16, '123', 'jaminan28-Mei-2017-19-59-246834929-sci-fi-wallpaper.jpg'),
(17, '123', 'jaminan28-Mei-2017-20-10-9121806.jpg'),
(18, '200000', 'jaminan9-Juni-2017-12-31-307dfd5bfa58d09c66a6a49251b882bcdd.jpg');

-- --------------------------------------------------------

--
-- Struktur dari tabel `tblpeminjaman`
--

CREATE TABLE `tblpeminjaman` (
  `id_peminjaman` int(11) NOT NULL,
  `no_ktp` varchar(255) NOT NULL,
  `tanggal_meminjam` date NOT NULL,
  `paket` varchar(100) NOT NULL,
  `hutang` double NOT NULL,
  `id_jaminan` int(11) NOT NULL,
  `lunas` tinyint(1) NOT NULL,
  `tanggal_lunas` date DEFAULT NULL,
  `nik` double NOT NULL,
  `besar_bunga` double NOT NULL COMMENT 'Besar Bunga dalam persen'
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `tblpeminjaman`
--

INSERT INTO `tblpeminjaman` (`id_peminjaman`, `no_ktp`, `tanggal_meminjam`, `paket`, `hutang`, `id_jaminan`, `lunas`, `tanggal_lunas`, `nik`, `besar_bunga`) VALUES
(1, '12151247', '2017-05-12', 'Pelunasan 4 Bulan', 120000, 10, 1, '2017-05-27', 12151247, 5),
(2, '1267421', '2017-05-28', 'Pelunasan 1 Bulan', 123, 17, 0, '0000-00-00', 12151247, 5),
(3, '12151247', '2017-06-09', 'Pelunasan 4 Bulan', 1000000, 18, 0, '0000-00-00', 12151247, 5);

-- --------------------------------------------------------

--
-- Struktur dari tabel `tblpengembalian`
--

CREATE TABLE `tblpengembalian` (
  `id_pengembalian` int(11) NOT NULL,
  `id_peminjaman` int(11) NOT NULL,
  `no_ktp` varchar(255) NOT NULL,
  `uang_bayar` double NOT NULL,
  `tanggal_bayar` date NOT NULL,
  `nik` double NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `tblpengembalian`
--

INSERT INTO `tblpengembalian` (`id_pengembalian`, `id_peminjaman`, `no_ktp`, `uang_bayar`, `tanggal_bayar`, `nik`) VALUES
(25, 1, '12151247', 120000, '2017-05-27', 12151247),
(26, 1, '12151247', 100000, '2017-06-01', 12151247),
(27, 1, '12151247', 8000, '2017-06-01', 12151247),
(29, 3, '12151247', 250000, '2017-06-09', 12151247),
(31, 3, '12151247', 250000, '2017-06-09', 12151247);

-- --------------------------------------------------------

--
-- Struktur dari tabel `tblpetugas`
--

CREATE TABLE `tblpetugas` (
  `nik` double NOT NULL,
  `password` varchar(100) NOT NULL,
  `nama` varchar(100) NOT NULL,
  `jabatan` varchar(20) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `tblpetugas`
--

INSERT INTO `tblpetugas` (`nik`, `password`, `nama`, `jabatan`) VALUES
(123, '123', 'Admin', 'admin'),
(12151244, '123', 'Elsa Dwi Isnawati', 'CO-Founder'),
(12151247, '123', 'Fajar Aziz Laksono', 'Founder & CEO'),
(12151375, '123', 'Esa Yuni Astri Diwati', 'Manager'),
(12151396, '123', 'Sofyan Afika Permata	', 'Producer'),
(12151413, '123', 'Imam Teguh Aji Saputro', 'Director');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `tblanggota`
--
ALTER TABLE `tblanggota`
  ADD PRIMARY KEY (`no_ktp`);

--
-- Indexes for table `tbljaminan`
--
ALTER TABLE `tbljaminan`
  ADD PRIMARY KEY (`id_jaminan`);

--
-- Indexes for table `tblpeminjaman`
--
ALTER TABLE `tblpeminjaman`
  ADD PRIMARY KEY (`id_peminjaman`),
  ADD KEY `no_ktp` (`no_ktp`),
  ADD KEY `id_jaminan` (`id_jaminan`),
  ADD KEY `nik` (`nik`),
  ADD KEY `no_ktp_2` (`no_ktp`),
  ADD KEY `no_ktp_3` (`no_ktp`),
  ADD KEY `id_jaminan_2` (`id_jaminan`);

--
-- Indexes for table `tblpengembalian`
--
ALTER TABLE `tblpengembalian`
  ADD PRIMARY KEY (`id_pengembalian`),
  ADD KEY `id_peminjaman` (`id_peminjaman`),
  ADD KEY `no_ktp` (`no_ktp`),
  ADD KEY `nik` (`nik`);

--
-- Indexes for table `tblpetugas`
--
ALTER TABLE `tblpetugas`
  ADD PRIMARY KEY (`nik`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `tbljaminan`
--
ALTER TABLE `tbljaminan`
  MODIFY `id_jaminan` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=19;
--
-- AUTO_INCREMENT for table `tblpeminjaman`
--
ALTER TABLE `tblpeminjaman`
  MODIFY `id_peminjaman` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=4;
--
-- AUTO_INCREMENT for table `tblpengembalian`
--
ALTER TABLE `tblpengembalian`
  MODIFY `id_pengembalian` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=32;
--
-- Ketidakleluasaan untuk tabel pelimpahan (Dumped Tables)
--

--
-- Ketidakleluasaan untuk tabel `tblpeminjaman`
--
ALTER TABLE `tblpeminjaman`
  ADD CONSTRAINT `tblpeminjaman_ibfk_4` FOREIGN KEY (`nik`) REFERENCES `tblpetugas` (`nik`),
  ADD CONSTRAINT `tblpeminjaman_ibfk_5` FOREIGN KEY (`id_jaminan`) REFERENCES `tbljaminan` (`id_jaminan`),
  ADD CONSTRAINT `tblpeminjaman_ibfk_6` FOREIGN KEY (`no_ktp`) REFERENCES `tblanggota` (`no_ktp`);

--
-- Ketidakleluasaan untuk tabel `tblpengembalian`
--
ALTER TABLE `tblpengembalian`
  ADD CONSTRAINT `tblpengembalian_ibfk_3` FOREIGN KEY (`nik`) REFERENCES `tblpetugas` (`nik`),
  ADD CONSTRAINT `tblpengembalian_ibfk_4` FOREIGN KEY (`id_peminjaman`) REFERENCES `tblpeminjaman` (`id_peminjaman`),
  ADD CONSTRAINT `tblpengembalian_ibfk_5` FOREIGN KEY (`no_ktp`) REFERENCES `tblanggota` (`no_ktp`);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
