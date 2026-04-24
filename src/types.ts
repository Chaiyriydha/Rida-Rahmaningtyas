export type SKPStatus = 'Belum Diambil' | 'Sudah Diambil';
export type SKPPeriode = 'Triwulan I' | 'Triwulan II' | 'Triwulan III' | 'Triwulan IV' | 'Tahunan';

export interface SKPRecord {
  id: string;
  nip: string;
  namaPegawai: string;
  jabatan: string;
  unitKerja: string;
  tahun: number;
  periode: SKPPeriode;
  jenisDokumen: string[];
  status: SKPStatus;
  fileUrl?: string;
  fileName?: string;
  pengambil?: string;
  unitKerjaPengambil?: string;
  tanggalAmbil?: string;
  keterangan?: string;
  ratingHasilKerja?: string;
  ratingHasilPerilaku?: string;
  predikatKinerja?: string;
}

export interface SKPStats {
  total: number;
  sudahDiambil: number;
  belumDiambil: number;
}

export type PLTStatus = 'Aktif' | 'Selesai';

export interface PLTRecord {
  id: string;
  nip: string;
  namaPegawai: string;
  jabatanAsli: string;
  jabatanPlt: string;
  unitKerja: string;
  unitKerjaTugas: string;
  noSk: string;
  tglMulai: string;
  tglSelesai?: string;
  tanggalSurat?: string;
  status: PLTStatus;
  keterangan?: string;
  pegawaiDigantikan?: string;
}

export interface PLHRecord {
  id: string;
  nip: string;
  namaPegawai: string;
  jabatanAsli: string;
  jabatanPlh: string;
  unitKerja: string;
  unitKerjaTugas: string;
  noSk: string;
  tglMulai: string;
  tglSelesai: string;
  tanggalSurat?: string;
  status: PLTStatus;
  keterangan?: string;
  pegawaiDigantikan?: string;
}

export interface TeamMember {
  nama: string;
  jabatan: string;
  status: string;
}

export interface WorkTeam {
  id: string;
  namaTim: string;
  unitKerja: string;
  ketua: TeamMember;
  anggota: TeamMember[];
  tahun: number;
  status: 'Aktif' | 'Non-Aktif';
  jenis: 'Tim Kerja' | 'Bagian';
  parentId?: string;
  createdAt?: any;
  updatedAt?: any;
  authorUid: string;
}
