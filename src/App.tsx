/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import * as React from 'react';
import { useState, useMemo } from 'react';
import { 
  Search, 
  Plus, 
  FileCheck, 
  FileClock, 
  Users, 
  LayoutDashboard,
  Filter,
  Download,
  MoreVertical,
  CheckCircle2,
  XCircle,
  Calendar as CalendarIcon,
  Building2,
  User,
  FileText,
  Eye,
  UploadCloud,
  X,
  UserCheck,
  UserCog
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { format } from 'date-fns';
import { id } from 'date-fns/locale';
import * as XLSX from 'xlsx';

import { Button } from '@/components/ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { 
  Table, 
  TableBody, 
  TableCell, 
  TableHead, 
  TableHeader, 
  TableRow 
} from '@/components/ui/table';
import { Badge } from '@/components/ui/badge';
import { 
  Dialog, 
  DialogContent, 
  DialogDescription, 
  DialogFooter, 
  DialogHeader, 
  DialogTitle,
  DialogTrigger,
  DialogClose
} from '@/components/ui/dialog';
import { Label } from '@/components/ui/label';
import { 
  Select, 
  SelectContent, 
  SelectItem, 
  SelectTrigger, 
  SelectValue 
} from '@/components/ui/select';
import { Checkbox } from '@/components/ui/checkbox';
import { Toaster } from '@/components/ui/sonner';
import { toast } from 'sonner';

import { SKPRecord, SKPStatus, SKPPeriode, PLTRecord, PLHRecord } from './types';
import { 
  MOCK_SKP_DATA, 
  MOCK_PLT_DATA,
  MOCK_PLH_DATA,
  UNIT_KERJA_LIST, 
  PERIODE_LIST, 
  JENIS_DOKUMEN_LIST,
  RATING_HASIL_KERJA_LIST,
  RATING_PERILAKU_LIST,
  PREDIKAT_KINERJA_LIST
} from './constants';
import { cn } from '@/lib/utils';

export default function App() {
  const [activeTab, setActiveTab] = useState<'skp' | 'plt-plh'>('skp');
  const [records, setRecords] = useState<SKPRecord[]>(MOCK_SKP_DATA);
  const [pltRecords, setPltRecords] = useState<PLTRecord[]>(MOCK_PLT_DATA);
  const [plhRecords, setPlhRecords] = useState<PLHRecord[]>(MOCK_PLH_DATA);
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [isAddDialogOpen, setIsAddDialogOpen] = useState(false);
  const [isTakeDialogOpen, setIsTakeDialogOpen] = useState(false);
  const [isDetailDialogOpen, setIsDetailDialogOpen] = useState(false);
  const [isEditDialogOpen, setIsEditDialogOpen] = useState(false);
  const [isViewFileOpen, setIsViewFileOpen] = useState(false);
  const [selectedRecord, setSelectedRecord] = useState<SKPRecord | null>(null);
  const [detailRecord, setDetailRecord] = useState<SKPRecord | null>(null);
  const [editingRecord, setEditingRecord] = useState<SKPRecord | null>(null);
  const [viewingFileUrl, setViewingFileUrl] = useState<string | null>(null);
  const [viewingFileName, setViewingFileName] = useState<string | null>(null);
  const [uploadPreview, setUploadPreview] = useState<string | null>(null);
  const [uploadFileName, setUploadFileName] = useState<string | null>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (uploadPreview) URL.revokeObjectURL(uploadPreview);
      setUploadPreview(URL.createObjectURL(file));
      setUploadFileName(file.name);
    } else {
      setUploadPreview(null);
      setUploadFileName(null);
    }
  };

  // Stats calculation
  const stats = useMemo(() => {
    const total = records.length;
    const sudahDiambil = records.filter(r => r.status === 'Sudah Diambil').length;
    const belumDiambil = total - sudahDiambil;
    return { total, sudahDiambil, belumDiambil };
  }, [records]);

  const pltPlhStats = useMemo(() => {
    const totalPlt = pltRecords.length;
    const totalPlh = plhRecords.length;
    const activePlt = pltRecords.filter(r => r.status === 'Aktif').length;
    const activePlh = plhRecords.filter(r => r.status === 'Aktif').length;
    return { totalPlt, totalPlh, activePlt, activePlh };
  }, [pltRecords, plhRecords]);

  const filteredPltRecords = useMemo(() => {
    return pltRecords.filter(record => 
      record.namaPegawai.toLowerCase().includes(searchQuery.toLowerCase()) ||
      record.nip.includes(searchQuery) ||
      record.jabatanPlt.toLowerCase().includes(searchQuery.toLowerCase()) ||
      record.unitKerja.toLowerCase().includes(searchQuery.toLowerCase())
    );
  }, [pltRecords, searchQuery]);

  const filteredPlhRecords = useMemo(() => {
    return plhRecords.filter(record => 
      record.namaPegawai.toLowerCase().includes(searchQuery.toLowerCase()) ||
      record.nip.includes(searchQuery) ||
      record.jabatanPlh.toLowerCase().includes(searchQuery.toLowerCase()) ||
      record.unitKerja.toLowerCase().includes(searchQuery.toLowerCase())
    );
  }, [plhRecords, searchQuery]);

  const getShortDocName = (name: string) => {
    const map: Record<string, string> = {
      'Sasaran Kinerja Pegawai': 'SKP',
      'Evaluasi Kinerja Pegawai': 'EKP',
      'Hasil Evaluasi Kinerja Pegawai Hasil Kerja Kuantitatif Tahunan': 'HEKPHKK',
    };
    return map[name] || name;
  };

  const filteredRecords = records.filter(record => {
    const matchesSearch = 
      record.namaPegawai.toLowerCase().includes(searchQuery.toLowerCase()) ||
      record.nip.includes(searchQuery) ||
      record.unitKerja.toLowerCase().includes(searchQuery.toLowerCase());
    
    const matchesStatus = statusFilter === 'all' || record.status === statusFilter;
    
    return matchesSearch && matchesStatus;
  });

  const handleAddRecord = (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const selectedJenisDokumen = JENIS_DOKUMEN_LIST.filter(item => formData.get(`jenisDokumen-${item}`) === 'on');
    
    const fileInput = e.currentTarget.querySelector('input[type="file"]') as HTMLInputElement;
    const file = fileInput?.files?.[0];
    
    let fileUrl = '';
    let fileName = '';
    
    if (file) {
      fileUrl = URL.createObjectURL(file);
      fileName = file.name;
    }

    const newRecord: SKPRecord = {
      id: Math.random().toString(36).substr(2, 9),
      nip: formData.get('nip') as string,
      namaPegawai: formData.get('namaPegawai') as string,
      jabatan: formData.get('jabatan') as string,
      unitKerja: formData.get('unitKerja') as string,
      tahun: parseInt(formData.get('tahun') as string),
      periode: formData.get('periode') as SKPPeriode,
      jenisDokumen: selectedJenisDokumen,
      status: 'Belum Diambil',
      ratingHasilKerja: formData.get('ratingHasilKerja') as string,
      ratingHasilPerilaku: formData.get('ratingHasilPerilaku') as string,
      predikatKinerja: formData.get('predikatKinerja') as string,
      fileUrl,
      fileName,
    };
    setRecords([newRecord, ...records]);
    setIsAddDialogOpen(false);
    setUploadPreview(null);
    setUploadFileName(null);
    toast.success('Data SKP berhasil ditambahkan');
  };

  const handleUpdateRecord = (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!editingRecord) return;

    const formData = new FormData(e.currentTarget);
    const selectedJenisDokumen = JENIS_DOKUMEN_LIST.filter(item => formData.get(`jenisDokumen-${item}`) === 'on');
    
    const fileInput = e.currentTarget.querySelector('input[type="file"]') as HTMLInputElement;
    const file = fileInput?.files?.[0];
    
    let fileUrl = editingRecord.fileUrl;
    let fileName = editingRecord.fileName;
    
    if (file) {
      fileUrl = URL.createObjectURL(file);
      fileName = file.name;
    }

    const updatedRecord: SKPRecord = {
      ...editingRecord,
      nip: formData.get('nip') as string,
      namaPegawai: formData.get('namaPegawai') as string,
      jabatan: formData.get('jabatan') as string,
      unitKerja: formData.get('unitKerja') as string,
      tahun: parseInt(formData.get('tahun') as string),
      periode: formData.get('periode') as SKPPeriode,
      jenisDokumen: selectedJenisDokumen,
      ratingHasilKerja: formData.get('ratingHasilKerja') as string,
      ratingHasilPerilaku: formData.get('ratingHasilPerilaku') as string,
      predikatKinerja: formData.get('predikatKinerja') as string,
      fileUrl,
      fileName,
    };

    setRecords(records.map(r => r.id === editingRecord.id ? updatedRecord : r));
    setIsEditDialogOpen(false);
    setEditingRecord(null);
    setUploadPreview(null);
    setUploadFileName(null);
    toast.success('Data SKP berhasil diperbarui');
  };

  const handleTakeSKP = (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!selectedRecord) return;

    const formData = new FormData(e.currentTarget);
    const tanggalAmbilInput = formData.get('tanggalAmbil') as string;
    
    const updatedRecords = records.map(r => {
      if (r.id === selectedRecord.id) {
        return {
          ...r,
          status: 'Sudah Diambil' as SKPStatus,
          pengambil: formData.get('pengambil') as string,
          unitKerjaPengambil: formData.get('unitKerjaPengambil') as string,
          tanggalAmbil: new Date(tanggalAmbilInput).toISOString(),
        };
      }
      return r;
    });

    setRecords(updatedRecords);
    setIsTakeDialogOpen(false);
    setSelectedRecord(null);
    toast.success('Status berkas berhasil diperbarui');
  };

  const handleExport = () => {
    if (records.length === 0) {
      toast.error('Tidak ada data untuk diekspor');
      return;
    }

    const exportData = records.map(record => ({
      'NIP': record.nip,
      'Nama Pegawai': record.namaPegawai,
      'Jabatan': record.jabatan,
      'Unit Kerja': record.unitKerja,
      'Tahun': record.tahun,
      'Periode': record.periode,
      'Jenis Dokumen': record.jenisDokumen.join(', '),
      'Status': record.status,
      'Rating Hasil Kerja': record.ratingHasilKerja || '-',
      'Rating Hasil Perilaku': record.ratingHasilPerilaku || '-',
      'Predikat Kinerja': record.predikatKinerja || '-',
      'Pengambil': record.pengambil || '-',
      'Unit Kerja Pengambil': record.unitKerjaPengambil || '-',
      'Tanggal Ambil': record.tanggalAmbil ? format(new Date(record.tanggalAmbil), 'dd MMMM yyyy, HH:mm', { locale: id }) : '-',
      'Keterangan': record.keterangan || '-',
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data SKP');
    
    // Auto-size columns
    const maxWidths = Object.keys(exportData[0]).map(key => {
      const headerLen = key.length;
      const maxDataLen = Math.max(...exportData.map(row => String(row[key as keyof typeof row]).length));
      return { wch: Math.max(headerLen, maxDataLen) + 2 };
    });
    worksheet['!cols'] = maxWidths;

    XLSX.writeFile(workbook, `Data_SKP_Export_${format(new Date(), 'yyyyMMdd_HHmm')}.xlsx`);
    toast.success('Data SKP berhasil diekspor ke Excel');
  };

  return (
    <div className="min-h-screen bg-[#F8F9FA] text-[#1A1A1A] font-sans selection:bg-primary/20">
      <Toaster position="top-right" />
      
      {/* Sidebar / Header */}
      <header className="sticky top-0 z-40 w-full border-b bg-white/80 backdrop-blur-md">
        <div className="container mx-auto flex h-16 items-center justify-between px-4">
          <div className="flex items-center gap-3">
            <div className="flex h-12 w-auto items-center justify-center overflow-hidden">
              <img 
                src="https://bskji.kemenperin.go.id/logo.png" 
                alt="Logo BSKJI" 
                className="h-full w-auto object-contain"
                referrerPolicy="no-referrer"
                onError={(e) => {
                  // Fallback if image fails to load
                  e.currentTarget.style.display = 'none';
                  e.currentTarget.parentElement!.innerHTML = `
                    <div class="flex h-10 w-10 items-center justify-center rounded-xl bg-primary text-primary-foreground shadow-lg shadow-primary/20">
                      <span class="font-bold text-lg">B</span>
                    </div>
                  `;
                }}
              />
            </div>
            <div className="hidden sm:block">
              <h1 className="text-lg font-bold tracking-tight text-primary">Monitoring SKP Pegawai</h1>
              <p className="text-[10px] text-muted-foreground font-medium uppercase tracking-wider leading-none">Badan Standardisasi dan Kebijakan Jasa Industri</p>
            </div>
          </div>

          <nav className="hidden md:flex items-center gap-1 bg-muted/30 p-1 rounded-lg">
            <Button 
              variant={activeTab === 'skp' ? 'default' : 'ghost'} 
              size="sm" 
              className={cn("h-8 gap-2 px-4 rounded-md transition-all", activeTab === 'skp' ? "shadow-sm" : "text-muted-foreground")}
              onClick={() => setActiveTab('skp')}
            >
              <FileText className="h-4 w-4" />
              SKP
            </Button>
            <Button 
              variant={activeTab === 'plt-plh' ? 'default' : 'ghost'} 
              size="sm" 
              className={cn("h-8 gap-2 px-4 rounded-md transition-all", activeTab === 'plt-plh' ? "shadow-sm" : "text-muted-foreground")}
              onClick={() => setActiveTab('plt-plh')}
            >
              <LayoutDashboard className="h-4 w-4" />
              PLT & PLH
            </Button>
          </nav>

          <div className="flex items-center gap-4">
            {activeTab === 'skp' && (
              <>
                <Button 
                  variant="outline" 
                  size="sm" 
                  className="hidden sm:flex gap-2"
                  onClick={handleExport}
                >
                  <Download className="h-4 w-4" />
                  Export Data
                </Button>
                <Dialog open={isAddDialogOpen} onOpenChange={(open) => {
                  setIsAddDialogOpen(open);
                  if (!open) {
                    if (uploadPreview) URL.revokeObjectURL(uploadPreview);
                    setUploadPreview(null);
                    setUploadFileName(null);
                  }
                }}>
                  <DialogTrigger 
                    nativeButton={true}
                    render={<Button size="sm" className="gap-2 shadow-md hover:shadow-lg transition-all" />}
                  >
                    <Plus className="h-4 w-4" />
                    Tambah SKP
                  </DialogTrigger>
                  <DialogContent className="sm:max-w-[550px] max-h-[90vh] overflow-y-auto">
                    <form onSubmit={handleAddRecord}>
                      <DialogHeader>
                        <DialogTitle>Tambah Data SKP Baru</DialogTitle>
                        <DialogDescription>
                          Masukkan informasi pegawai untuk monitoring berkas SKP.
                        </DialogDescription>
                      </DialogHeader>
                      <div className="grid gap-4 py-4">
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                          <div className="grid gap-2">
                            <Label htmlFor="namaPegawai">Nama Lengkap</Label>
                            <Input id="namaPegawai" name="namaPegawai" placeholder="Contoh: Ahmad Subarjo" required />
                          </div>
                          <div className="grid gap-2">
                            <Label htmlFor="nip">NIP</Label>
                            <Input id="nip" name="nip" placeholder="18 digit NIP" required />
                          </div>
                        </div>
                        
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                          <div className="grid gap-2">
                            <Label htmlFor="jabatan">Jabatan</Label>
                            <Input id="jabatan" name="jabatan" placeholder="Contoh: Analis Kepegawaian" required />
                          </div>
                          <div className="grid gap-2">
                            <Label htmlFor="unitKerja">Unit Kerja</Label>
                            <Select name="unitKerja" required>
                              <SelectTrigger>
                                <SelectValue placeholder="Pilih Unit Kerja" />
                              </SelectTrigger>
                              <SelectContent>
                                {UNIT_KERJA_LIST.map(unit => (
                                  <SelectItem key={unit} value={unit}>{unit}</SelectItem>
                                ))}
                              </SelectContent>
                            </Select>
                          </div>
                        </div>

                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                          <div className="grid gap-2">
                            <Label htmlFor="tahun">Tahun SKP</Label>
                            <Input id="tahun" name="tahun" type="number" defaultValue={new Date().getFullYear()} required />
                          </div>
                          <div className="grid gap-2">
                            <Label htmlFor="periode">Periode SKP</Label>
                            <Select name="periode" defaultValue="Tahunan" required>
                              <SelectTrigger>
                                <SelectValue placeholder="Pilih Periode" />
                              </SelectTrigger>
                              <SelectContent>
                                {PERIODE_LIST.map(periode => (
                                  <SelectItem key={periode} value={periode}>{periode}</SelectItem>
                                ))}
                              </SelectContent>
                            </Select>
                          </div>
                        </div>

                        <div className="grid gap-3">
                          <Label>Jenis Dokumen</Label>
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 border rounded-lg p-3 bg-muted/20">
                            {JENIS_DOKUMEN_LIST.map((item) => (
                              <div key={item} className="flex items-center space-x-2">
                                <Checkbox id={`jenisDokumen-${item}`} name={`jenisDokumen-${item}`} />
                                <label
                                  htmlFor={`jenisDokumen-${item}`}
                                  className="text-[11px] font-medium leading-none peer-disabled:cursor-not-allowed peer-disabled:opacity-70 cursor-pointer"
                                >
                                  {item}
                                </label>
                              </div>
                            ))}
                          </div>
                        </div>

                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                          <div className="grid gap-2">
                            <Label htmlFor="ratingHasilKerja">Rating Hasil Kerja</Label>
                            <Select name="ratingHasilKerja">
                              <SelectTrigger>
                                <SelectValue placeholder="Pilih Rating" />
                              </SelectTrigger>
                              <SelectContent>
                                {RATING_HASIL_KERJA_LIST.map(rating => (
                                  <SelectItem key={rating} value={rating}>{rating}</SelectItem>
                                ))}
                              </SelectContent>
                            </Select>
                          </div>
                          <div className="grid gap-2">
                            <Label htmlFor="ratingHasilPerilaku">Rating Hasil Perilaku</Label>
                            <Select name="ratingHasilPerilaku">
                              <SelectTrigger>
                                <SelectValue placeholder="Pilih Rating" />
                              </SelectTrigger>
                              <SelectContent>
                                {RATING_PERILAKU_LIST.map(rating => (
                                  <SelectItem key={rating} value={rating}>{rating}</SelectItem>
                                ))}
                              </SelectContent>
                            </Select>
                          </div>
                        </div>

                        <div className="grid gap-2">
                          <Label htmlFor="predikatKinerja">Predikat Kinerja</Label>
                          <Select name="predikatKinerja">
                            <SelectTrigger>
                              <SelectValue placeholder="Pilih Predikat" />
                            </SelectTrigger>
                            <SelectContent>
                              {PREDIKAT_KINERJA_LIST.map(predikat => (
                                <SelectItem key={predikat} value={predikat}>{predikat}</SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>

                        <div className="grid gap-2">
                          <Label htmlFor="file">Upload Berkas SKP (PDF/Gambar)</Label>
                          <div className="flex flex-col gap-3">
                            <div className="flex items-center justify-center w-full">
                              <label className="flex flex-col items-center justify-center w-full h-24 border-2 border-dashed rounded-lg cursor-pointer bg-muted/50 hover:bg-muted transition-colors border-muted-foreground/25">
                                <div className="flex flex-col items-center justify-center pt-2 pb-3">
                                  <UploadCloud className="w-6 h-6 mb-2 text-muted-foreground" />
                                  <p className="text-xs text-muted-foreground">
                                    <span className="font-semibold">Klik untuk upload</span> atau drag and drop
                                  </p>
                                </div>
                                <input 
                                  id="file" 
                                  name="file" 
                                  type="file" 
                                  className="hidden" 
                                  accept=".pdf,image/*" 
                                  onChange={handleFileChange}
                                />
                              </label>
                            </div>
                            
                            {uploadPreview && (
                              <div className="p-3 rounded-lg border bg-muted/30 flex items-center gap-3">
                                <div className="h-12 w-12 rounded border bg-background flex items-center justify-center overflow-hidden shrink-0">
                                  {uploadFileName?.toLowerCase().endsWith('.pdf') ? (
                                    <FileText className="h-6 w-6 text-red-500" />
                                  ) : (
                                    <img 
                                      src={uploadPreview} 
                                      alt="Preview" 
                                      className="h-full w-full object-cover"
                                      referrerPolicy="no-referrer"
                                    />
                                  )}
                                </div>
                                <div className="flex-1 min-w-0">
                                  <p className="text-xs font-medium truncate">{uploadFileName}</p>
                                  <p className="text-[10px] text-muted-foreground uppercase">Siap untuk disimpan</p>
                                </div>
                                <Button 
                                  type="button" 
                                  variant="ghost" 
                                  size="icon-xs" 
                                  className="text-muted-foreground hover:text-destructive"
                                  onClick={() => {
                                    if (uploadPreview) URL.revokeObjectURL(uploadPreview);
                                    setUploadPreview(null);
                                    setUploadFileName(null);
                                    // Reset file input
                                    const fileInput = document.getElementById('file') as HTMLInputElement;
                                    if (fileInput) fileInput.value = '';
                                  }}
                                >
                                  <X className="h-3.5 w-3.5" />
                                </Button>
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                      <DialogFooter className="gap-2 sm:gap-0 sticky bottom-0 bg-background pt-2 border-t mt-2">
                        <DialogClose 
                          nativeButton={true}
                          render={<Button type="button" variant="outline">Batal</Button>} 
                        />
                        <Button type="submit" className="bg-primary hover:bg-primary/90">
                          Simpan Data SKP
                        </Button>
                      </DialogFooter>
                    </form>
                  </DialogContent>
                </Dialog>
              </>
            )}
            {activeTab === 'plt-plh' && (
              <Button 
                size="sm" 
                className="gap-2 shadow-md hover:shadow-lg transition-all"
                onClick={() => toast.info('Fitur Tambah PLT/PLH akan segera hadir')}
              >
                <Plus className="h-4 w-4" />
                Tambah PLT/PLH
              </Button>
            )}

            {/* Edit SKP Dialog */}
            <Dialog open={isEditDialogOpen} onOpenChange={(open) => {
              setIsEditDialogOpen(open);
              if (!open) {
                setEditingRecord(null);
                if (uploadPreview) URL.revokeObjectURL(uploadPreview);
                setUploadPreview(null);
                setUploadFileName(null);
              }
            }}>
              <DialogContent className="sm:max-w-[550px] max-h-[90vh] overflow-y-auto">
                {editingRecord && (
                  <form onSubmit={handleUpdateRecord}>
                    <DialogHeader>
                      <DialogTitle>Edit Data SKP</DialogTitle>
                      <DialogDescription>
                        Perbarui informasi pegawai jika terjadi mutasi atau perubahan data lainnya.
                      </DialogDescription>
                    </DialogHeader>
                    <div className="grid gap-4 py-4">
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div className="grid gap-2">
                          <Label htmlFor="edit-namaPegawai">Nama Lengkap</Label>
                          <Input id="edit-namaPegawai" name="namaPegawai" defaultValue={editingRecord.namaPegawai} required />
                        </div>
                        <div className="grid gap-2">
                          <Label htmlFor="edit-nip">NIP</Label>
                          <Input id="edit-nip" name="nip" defaultValue={editingRecord.nip} required />
                        </div>
                      </div>
                      
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div className="grid gap-2">
                          <Label htmlFor="edit-jabatan">Jabatan</Label>
                          <Input id="edit-jabatan" name="jabatan" defaultValue={editingRecord.jabatan} required />
                        </div>
                        <div className="grid gap-2">
                          <Label htmlFor="edit-unitKerja">Unit Kerja</Label>
                          <Select name="unitKerja" defaultValue={editingRecord.unitKerja} required>
                            <SelectTrigger>
                              <SelectValue placeholder="Pilih Unit Kerja" />
                            </SelectTrigger>
                            <SelectContent>
                              {UNIT_KERJA_LIST.map(unit => (
                                <SelectItem key={unit} value={unit}>{unit}</SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div className="grid gap-2">
                          <Label htmlFor="edit-tahun">Tahun SKP</Label>
                          <Input id="edit-tahun" name="tahun" type="number" defaultValue={editingRecord.tahun} required />
                        </div>
                        <div className="grid gap-2">
                          <Label htmlFor="edit-periode">Periode SKP</Label>
                          <Select name="periode" defaultValue={editingRecord.periode} required>
                            <SelectTrigger>
                              <SelectValue placeholder="Pilih Periode" />
                            </SelectTrigger>
                            <SelectContent>
                              {PERIODE_LIST.map(periode => (
                                <SelectItem key={periode} value={periode}>{periode}</SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>
                      </div>

                      <div className="grid gap-3">
                        <Label>Jenis Dokumen</Label>
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 border rounded-lg p-3 bg-muted/20">
                          {JENIS_DOKUMEN_LIST.map((item) => (
                            <div key={item} className="flex items-center space-x-2">
                              <Checkbox 
                                id={`edit-jenisDokumen-${item}`} 
                                name={`jenisDokumen-${item}`} 
                                defaultChecked={editingRecord.jenisDokumen.includes(item)}
                              />
                              <label
                                htmlFor={`edit-jenisDokumen-${item}`}
                                className="text-[11px] font-medium leading-none peer-disabled:cursor-not-allowed peer-disabled:opacity-70 cursor-pointer"
                              >
                                {item}
                              </label>
                            </div>
                          ))}
                        </div>
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div className="grid gap-2">
                          <Label htmlFor="edit-ratingHasilKerja">Rating Hasil Kerja</Label>
                          <Select name="ratingHasilKerja" defaultValue={editingRecord.ratingHasilKerja}>
                            <SelectTrigger>
                              <SelectValue placeholder="Pilih Rating" />
                            </SelectTrigger>
                            <SelectContent>
                              {RATING_HASIL_KERJA_LIST.map(rating => (
                                <SelectItem key={rating} value={rating}>{rating}</SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>
                        <div className="grid gap-2">
                          <Label htmlFor="edit-ratingHasilPerilaku">Rating Hasil Perilaku</Label>
                          <Select name="ratingHasilPerilaku" defaultValue={editingRecord.ratingHasilPerilaku}>
                            <SelectTrigger>
                              <SelectValue placeholder="Pilih Rating" />
                            </SelectTrigger>
                            <SelectContent>
                              {RATING_PERILAKU_LIST.map(rating => (
                                <SelectItem key={rating} value={rating}>{rating}</SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>
                      </div>

                      <div className="grid gap-2">
                        <Label htmlFor="edit-predikatKinerja">Predikat Kinerja</Label>
                        <Select name="predikatKinerja" defaultValue={editingRecord.predikatKinerja}>
                          <SelectTrigger>
                            <SelectValue placeholder="Pilih Predikat" />
                          </SelectTrigger>
                          <SelectContent>
                            {PREDIKAT_KINERJA_LIST.map(predikat => (
                              <SelectItem key={predikat} value={predikat}>{predikat}</SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </div>

                      <div className="grid gap-2">
                        <Label htmlFor="edit-file">Ganti Berkas SKP (Opsional)</Label>
                        <div className="flex flex-col gap-3">
                          <div className="flex items-center justify-center w-full">
                            <label className="flex flex-col items-center justify-center w-full h-24 border-2 border-dashed rounded-lg cursor-pointer bg-muted/50 hover:bg-muted transition-colors border-muted-foreground/25">
                              <div className="flex flex-col items-center justify-center pt-2 pb-3">
                                <UploadCloud className="w-6 h-6 mb-2 text-muted-foreground" />
                                <p className="text-xs text-muted-foreground">
                                  <span className="font-semibold">Klik untuk ganti berkas</span> atau drag and drop
                                </p>
                              </div>
                              <input 
                                id="edit-file" 
                                name="file" 
                                type="file" 
                                className="hidden" 
                                accept=".pdf,image/*" 
                                onChange={handleFileChange}
                              />
                            </label>
                          </div>
                          
                          {uploadPreview ? (
                            <div className="p-3 rounded-lg border bg-muted/30 flex items-center gap-3">
                              <div className="h-12 w-12 rounded border bg-background flex items-center justify-center overflow-hidden shrink-0">
                                {uploadFileName?.toLowerCase().endsWith('.pdf') ? (
                                  <FileText className="h-6 w-6 text-red-500" />
                                ) : (
                                  <img 
                                    src={uploadPreview} 
                                    alt="Preview" 
                                    className="h-full w-full object-cover"
                                    referrerPolicy="no-referrer"
                                  />
                                )}
                              </div>
                              <div className="flex-1 min-w-0">
                                <p className="text-xs font-medium truncate">{uploadFileName}</p>
                                <p className="text-[10px] text-muted-foreground uppercase">Berkas Baru Terpilih</p>
                              </div>
                              <Button 
                                type="button" 
                                variant="ghost" 
                                size="icon-xs" 
                                className="text-muted-foreground hover:text-destructive"
                                onClick={() => {
                                  if (uploadPreview) URL.revokeObjectURL(uploadPreview);
                                  setUploadPreview(null);
                                  setUploadFileName(null);
                                  const fileInput = document.getElementById('edit-file') as HTMLInputElement;
                                  if (fileInput) fileInput.value = '';
                                }}
                              >
                                <X className="h-3.5 w-3.5" />
                              </Button>
                            </div>
                          ) : editingRecord.fileName && (
                            <div className="p-3 rounded-lg border bg-muted/10 flex items-center gap-3">
                              <div className="h-10 w-10 rounded border bg-background flex items-center justify-center shrink-0">
                                <FileText className="h-5 w-5 text-muted-foreground" />
                              </div>
                              <div className="flex-1 min-w-0">
                                <p className="text-xs font-medium truncate">{editingRecord.fileName}</p>
                                <p className="text-[10px] text-muted-foreground uppercase">Berkas Saat Ini</p>
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                    <DialogFooter className="gap-2 sm:gap-0 sticky bottom-0 bg-background pt-2 border-t mt-2">
                      <DialogClose 
                        nativeButton={true}
                        render={<Button type="button" variant="outline">Batal</Button>} 
                      />
                      <Button type="submit" className="bg-primary hover:bg-primary/90">
                        Simpan Perubahan
                      </Button>
                    </DialogFooter>
                  </form>
                )}
              </DialogContent>
            </Dialog>
          </div>
        </div>
      </header>

      <main className="container mx-auto p-4 md:p-8 space-y-8">
        {activeTab === 'skp' ? (
          <>
            {/* Stats Grid */}
            <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3">
              <motion.div
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.1 }}
              >
                <Card className="overflow-hidden border-none shadow-sm hover:shadow-md transition-shadow">
                  <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">Total Berkas SKP</CardTitle>
                    <Users className="h-4 w-4 text-primary" />
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold">{stats.total}</div>
                    <p className="text-xs text-muted-foreground mt-1">Seluruh data pegawai terdaftar</p>
                  </CardContent>
                </Card>
              </motion.div>

              <motion.div
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.2 }}
              >
                <Card className="overflow-hidden border-none shadow-sm hover:shadow-md transition-shadow">
                  <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">Sudah Diambil</CardTitle>
                    <FileCheck className="h-4 w-4 text-emerald-500" />
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold text-emerald-600">{stats.sudahDiambil}</div>
                    <div className="flex items-center gap-2 mt-1">
                      <div className="h-1.5 w-full rounded-full bg-muted overflow-hidden">
                        <div 
                          className="h-full bg-emerald-500 transition-all duration-500" 
                          style={{ width: `${(stats.sudahDiambil / stats.total) * 100}%` }}
                        />
                      </div>
                      <span className="text-[10px] font-bold text-emerald-600">
                        {Math.round((stats.sudahDiambil / stats.total) * 100)}%
                      </span>
                    </div>
                  </CardContent>
                </Card>
              </motion.div>

              <motion.div
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.3 }}
              >
                <Card className="overflow-hidden border-none shadow-sm hover:shadow-md transition-shadow">
                  <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">Belum Diambil</CardTitle>
                    <FileClock className="h-4 w-4 text-amber-500" />
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold text-amber-600">{stats.belumDiambil}</div>
                    <p className="text-xs text-muted-foreground mt-1">Menunggu proses pengambilan</p>
                  </CardContent>
                </Card>
              </motion.div>
            </div>

            {/* Filters & Table */}
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-white border-b pb-6">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div>
                    <CardTitle>Daftar Berkas SKP</CardTitle>
                    <CardDescription>Kelola dan pantau status pengambilan berkas SKP pegawai.</CardDescription>
                  </div>
                  <div className="flex flex-wrap items-center gap-3">
                    <div className="relative w-full md:w-64">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
                      <Input 
                        placeholder="Cari Nama, NIP, atau Unit..." 
                        className="pl-9 bg-muted/50 border-none focus-visible:ring-1"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                      />
                    </div>
                    <Select value={statusFilter} onValueChange={setStatusFilter}>
                      <SelectTrigger className="w-[160px] bg-muted/50 border-none">
                        <Filter className="h-4 w-4 mr-2 opacity-50" />
                        <SelectValue placeholder="Filter Status" />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="all">Semua Status</SelectItem>
                        <SelectItem value="Belum Diambil">Belum Diambil</SelectItem>
                        <SelectItem value="Sudah Diambil">Sudah Diambil</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                </div>
              </CardHeader>
              <CardContent className="p-0">
                <div className="overflow-x-auto">
                  <Table>
                    <TableHeader className="bg-muted/30">
                      <TableRow className="hover:bg-transparent">
                        <TableHead className="w-[120px] font-bold text-xs uppercase tracking-wider">Tahun & Periode</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Pegawai</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Unit Kerja</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Status</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Keterangan Pengambilan</TableHead>
                        <TableHead className="text-right font-bold text-xs uppercase tracking-wider">Aksi</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      <AnimatePresence mode="popLayout">
                        {filteredRecords.length > 0 ? (
                          filteredRecords.map((record) => (
                            <motion.tr
                              key={record.id}
                              layout
                              initial={{ opacity: 0 }}
                              animate={{ opacity: 1 }}
                              exit={{ opacity: 0 }}
                              className="group hover:bg-muted/20 transition-colors"
                            >
                              <TableCell>
                                <div className="flex flex-col">
                                  <span className="font-mono text-sm font-medium">{record.tahun}</span>
                                  <span className="text-[10px] text-muted-foreground font-bold uppercase">{record.periode}</span>
                                </div>
                              </TableCell>
                              <TableCell>
                                <div className="flex flex-col">
                                  <button 
                                    onClick={() => {
                                      setDetailRecord(record);
                                      setIsDetailDialogOpen(true);
                                    }}
                                    className="text-left hover:text-primary hover:underline transition-colors"
                                  >
                                    <span className="font-semibold text-sm">{record.namaPegawai}</span>
                                  </button>
                                  <span className="text-[10px] text-muted-foreground font-medium">{record.jabatan}</span>
                                  <span className="text-[10px] text-muted-foreground font-mono">{record.nip}</span>
                                </div>
                              </TableCell>
                              <TableCell>
                                <div className="flex items-center gap-2 text-sm">
                                  <Building2 className="h-3.5 w-3.5 text-muted-foreground" />
                                  {record.unitKerja}
                                </div>
                              </TableCell>
                              <TableCell>
                                <Badge 
                                  variant={record.status === 'Sudah Diambil' ? 'default' : 'secondary'}
                                  className={cn(
                                    "rounded-full px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider",
                                    record.status === 'Sudah Diambil' 
                                      ? "bg-emerald-100 text-emerald-700 hover:bg-emerald-100" 
                                      : "bg-amber-100 text-amber-700 hover:bg-amber-100"
                                  )}
                                >
                                  {record.status === 'Sudah Diambil' ? (
                                    <CheckCircle2 className="h-3 w-3 mr-1" />
                                  ) : (
                                    <XCircle className="h-3 w-3 mr-1" />
                                  )}
                                  {record.status}
                                </Badge>
                              </TableCell>
                              <TableCell>
                                {record.status === 'Sudah Diambil' ? (
                                  <div className="flex flex-col gap-1">
                                    <div className="flex items-center gap-1.5 text-xs font-medium">
                                      <User className="h-3 w-3 text-muted-foreground" />
                                      {record.pengambil}
                                    </div>
                                    <div className="flex items-center gap-1.5 text-[10px] text-muted-foreground">
                                      <CalendarIcon className="h-3 w-3" />
                                      {record.tanggalAmbil && format(new Date(record.tanggalAmbil), 'dd MMM yyyy, HH:mm', { locale: id })}
                                    </div>
                                    <div className="text-[10px] text-muted-foreground italic">
                                      Unit: {record.unitKerjaPengambil}
                                    </div>
                                  </div>
                                ) : (
                                  <span className="text-xs text-muted-foreground italic">Belum ada data</span>
                                )}
                              </TableCell>
                              <TableCell className="text-right">
                                {record.status === 'Belum Diambil' ? (
                                  <Button 
                                    size="sm" 
                                    variant="outline"
                                    className="h-8 text-xs gap-1.5 border-emerald-200 text-emerald-700 hover:bg-emerald-50 hover:text-emerald-800"
                                    onClick={() => {
                                      setSelectedRecord(record);
                                      setIsTakeDialogOpen(true);
                                    }}
                                  >
                                    <FileCheck className="h-3.5 w-3.5" />
                                    Ambil Berkas
                                  </Button>
                                ) : (
                                  <Button size="icon" variant="ghost" className="h-8 w-8 opacity-0 group-hover:opacity-100 transition-opacity">
                                    <MoreVertical className="h-4 w-4" />
                                  </Button>
                                )}
                              </TableCell>
                            </motion.tr>
                          ))
                        ) : (
                          <TableRow>
                            <TableCell colSpan={6} className="h-32 text-center text-muted-foreground">
                              Tidak ada data yang ditemukan.
                            </TableCell>
                          </TableRow>
                        )}
                      </AnimatePresence>
                    </TableBody>
                  </Table>
                </div>
              </CardContent>
            </Card>
          </>
        ) : (
          <div className="space-y-8">
            {/* PLT/PLH Stats Grid */}
            <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4">
              <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.1 }}>
                <Card className="overflow-hidden border-none shadow-sm hover:shadow-md transition-shadow">
                  <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">Total PLT</CardTitle>
                    <UserCog className="h-4 w-4 text-primary" />
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold">{pltPlhStats.totalPlt}</div>
                    <p className="text-xs text-muted-foreground mt-1">Pelaksana Tugas terdaftar</p>
                  </CardContent>
                </Card>
              </motion.div>

              <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.2 }}>
                <Card className="overflow-hidden border-none shadow-sm hover:shadow-md transition-shadow">
                  <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">PLT Aktif</CardTitle>
                    <CheckCircle2 className="h-4 w-4 text-emerald-500" />
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold text-emerald-600">{pltPlhStats.activePlt}</div>
                    <p className="text-xs text-muted-foreground mt-1">Sedang menjabat</p>
                  </CardContent>
                </Card>
              </motion.div>

              <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.3 }}>
                <Card className="overflow-hidden border-none shadow-sm hover:shadow-md transition-shadow">
                  <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">Total PLH</CardTitle>
                    <UserCheck className="h-4 w-4 text-blue-500" />
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold text-blue-600">{pltPlhStats.totalPlh}</div>
                    <p className="text-xs text-muted-foreground mt-1">Pelaksana Harian terdaftar</p>
                  </CardContent>
                </Card>
              </motion.div>

              <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.4 }}>
                <Card className="overflow-hidden border-none shadow-sm hover:shadow-md transition-shadow">
                  <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">PLH Aktif</CardTitle>
                    <CheckCircle2 className="h-4 w-4 text-blue-400" />
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold text-blue-500">{pltPlhStats.activePlh}</div>
                    <p className="text-xs text-muted-foreground mt-1">Sedang menjabat</p>
                  </CardContent>
                </Card>
              </motion.div>
            </div>

            {/* PLT Table */}
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-white border-b pb-6">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div>
                    <CardTitle>Daftar Pelaksana Tugas (PLT)</CardTitle>
                    <CardDescription>Monitoring penugasan Pelaksana Tugas di lingkungan BSKJI.</CardDescription>
                  </div>
                  <div className="relative w-full md:w-64">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
                    <Input 
                      placeholder="Cari PLT..." 
                      className="pl-9 bg-muted/50 border-none focus-visible:ring-1"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                    />
                  </div>
                </div>
              </CardHeader>
              <CardContent className="p-0">
                <div className="overflow-x-auto">
                  <Table>
                    <TableHeader className="bg-muted/30">
                      <TableRow className="hover:bg-transparent">
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Pegawai</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Jabatan PLT</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Unit Kerja</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Periode</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Status</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {filteredPltRecords.length > 0 ? (
                        filteredPltRecords.map((record) => (
                          <TableRow key={record.id} className="group hover:bg-muted/20 transition-colors">
                            <TableCell>
                              <div className="flex flex-col">
                                <span className="font-semibold text-sm">{record.namaPegawai}</span>
                                <span className="text-[10px] text-muted-foreground font-mono">{record.nip}</span>
                                <span className="text-[10px] text-muted-foreground italic">Asli: {record.jabatanAsli}</span>
                              </div>
                            </TableCell>
                            <TableCell>
                              <div className="flex flex-col">
                                <span className="font-medium text-sm text-primary">{record.jabatanPlt}</span>
                                <span className="text-[10px] text-muted-foreground font-medium">SK: {record.noSk}</span>
                              </div>
                            </TableCell>
                            <TableCell>
                              <div className="flex items-center gap-2 text-sm">
                                <Building2 className="h-3.5 w-3.5 text-muted-foreground" />
                                {record.unitKerja}
                              </div>
                            </TableCell>
                            <TableCell>
                              <div className="flex flex-col text-xs">
                                <span>{format(new Date(record.tglMulai), 'dd MMM yyyy')}</span>
                                <span className="text-muted-foreground">s.d. {format(new Date(record.tglSelesai), 'dd MMM yyyy')}</span>
                              </div>
                            </TableCell>
                            <TableCell>
                              <Badge 
                                variant={record.status === 'Aktif' ? 'default' : 'secondary'}
                                className={cn(
                                  "rounded-full px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider",
                                  record.status === 'Aktif' 
                                    ? "bg-emerald-100 text-emerald-700" 
                                    : "bg-muted text-muted-foreground"
                                )}
                              >
                                {record.status}
                              </Badge>
                            </TableCell>
                          </TableRow>
                        ))
                      ) : (
                        <TableRow>
                          <TableCell colSpan={5} className="h-24 text-center text-muted-foreground">Tidak ada data PLT.</TableCell>
                        </TableRow>
                      )}
                    </TableBody>
                  </Table>
                </div>
              </CardContent>
            </Card>

            {/* PLH Table */}
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-white border-b pb-6">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div>
                    <CardTitle>Daftar Pelaksana Harian (PLH)</CardTitle>
                    <CardDescription>Monitoring penugasan Pelaksana Harian di lingkungan BSKJI.</CardDescription>
                  </div>
                </div>
              </CardHeader>
              <CardContent className="p-0">
                <div className="overflow-x-auto">
                  <Table>
                    <TableHeader className="bg-muted/30">
                      <TableRow className="hover:bg-transparent">
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Pegawai</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Jabatan PLH</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Unit Kerja</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Periode</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Status</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {filteredPlhRecords.length > 0 ? (
                        filteredPlhRecords.map((record) => (
                          <TableRow key={record.id} className="group hover:bg-muted/20 transition-colors">
                            <TableCell>
                              <div className="flex flex-col">
                                <span className="font-semibold text-sm">{record.namaPegawai}</span>
                                <span className="text-[10px] text-muted-foreground font-mono">{record.nip}</span>
                                <span className="text-[10px] text-muted-foreground italic">Asli: {record.jabatanAsli}</span>
                              </div>
                            </TableCell>
                            <TableCell>
                              <div className="flex flex-col">
                                <span className="font-medium text-sm text-blue-600">{record.jabatanPlh}</span>
                                <span className="text-[10px] text-muted-foreground font-medium">SK: {record.noSk}</span>
                              </div>
                            </TableCell>
                            <TableCell>
                              <div className="flex items-center gap-2 text-sm">
                                <Building2 className="h-3.5 w-3.5 text-muted-foreground" />
                                {record.unitKerja}
                              </div>
                            </TableCell>
                            <TableCell>
                              <div className="flex flex-col text-xs">
                                <span>{format(new Date(record.tglMulai), 'dd MMM yyyy')}</span>
                                <span className="text-muted-foreground">s.d. {format(new Date(record.tglSelesai), 'dd MMM yyyy')}</span>
                              </div>
                            </TableCell>
                            <TableCell>
                              <Badge 
                                variant={record.status === 'Aktif' ? 'default' : 'secondary'}
                                className={cn(
                                  "rounded-full px-3 py-0.5 text-[10px] font-bold uppercase tracking-wider",
                                  record.status === 'Aktif' 
                                    ? "bg-blue-100 text-blue-700" 
                                    : "bg-muted text-muted-foreground"
                                )}
                              >
                                {record.status}
                              </Badge>
                            </TableCell>
                          </TableRow>
                        ))
                      ) : (
                        <TableRow>
                          <TableCell colSpan={5} className="h-24 text-center text-muted-foreground">Tidak ada data PLH.</TableCell>
                        </TableRow>
                      )}
                    </TableBody>
                  </Table>
                </div>
              </CardContent>
            </Card>
          </div>
        )}
      </main>

      {/* Take SKP Dialog */}
      <Dialog open={isTakeDialogOpen} onOpenChange={setIsTakeDialogOpen}>
        <DialogContent className="sm:max-w-[425px] max-h-[90vh] overflow-y-auto">
          <form onSubmit={handleTakeSKP}>
            <DialogHeader>
              <DialogTitle>Konfirmasi Pengambilan Berkas</DialogTitle>
              <DialogDescription>
                Catat informasi siapa yang mengambil berkas SKP milik <strong>{selectedRecord?.namaPegawai}</strong>.
              </DialogDescription>
            </DialogHeader>
            <div className="grid gap-4 py-4">
              <div className="grid gap-2">
                <Label htmlFor="pengambil">Nama Pengambil</Label>
                <Input id="pengambil" name="pengambil" placeholder="Nama lengkap pengambil" required />
              </div>
              <div className="grid gap-2">
                <Label htmlFor="unitKerjaPengambil">Unit Kerja Pengambil</Label>
                <Select name="unitKerjaPengambil" required>
                  <SelectTrigger>
                    <SelectValue placeholder="Pilih Unit Kerja" />
                  </SelectTrigger>
                  <SelectContent>
                    {UNIT_KERJA_LIST.map(unit => (
                      <SelectItem key={unit} value={unit}>{unit}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div className="grid gap-2">
                <Label htmlFor="tanggalAmbil">Tanggal Pengambilan</Label>
                <Input 
                  id="tanggalAmbil" 
                  name="tanggalAmbil" 
                  type="datetime-local" 
                  defaultValue={new Date().toISOString().slice(0, 16)} 
                  required 
                />
              </div>
              <div className="p-3 rounded-lg bg-muted/50 border text-[11px] text-muted-foreground">
                <p><strong>Catatan:</strong> Anda dapat menyesuaikan tanggal dan waktu pengambilan jika diperlukan.</p>
              </div>
            </div>
            <DialogFooter>
              <DialogClose 
                nativeButton={true}
                render={<Button type="button" variant="ghost" />}
              >
                Batal
              </DialogClose>
              <Button type="submit" className="bg-emerald-600 hover:bg-emerald-700">Konfirmasi Ambil</Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>

      {/* Detail Dialog */}
      <Dialog open={isDetailDialogOpen} onOpenChange={setIsDetailDialogOpen}>
        <DialogContent className="sm:max-w-[500px] max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <User className="h-5 w-5 text-primary" />
              Detail Informasi Pegawai
            </DialogTitle>
            <DialogDescription>
              Informasi lengkap terkait berkas SKP dan data kepegawaian.
            </DialogDescription>
          </DialogHeader>
          {detailRecord && (
            <div className="grid gap-6 py-4">
              <div className="flex flex-col items-center gap-2 p-4 bg-muted/30 rounded-xl border border-muted-foreground/10">
                <div className="h-16 w-16 rounded-full bg-primary/10 flex items-center justify-center">
                  <User className="h-8 w-8 text-primary" />
                </div>
                <div className="text-center">
                  <h3 className="font-bold text-lg">{detailRecord.namaPegawai}</h3>
                  <p className="text-sm text-muted-foreground font-mono">{detailRecord.nip}</p>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <Label className="text-[10px] uppercase tracking-wider text-muted-foreground">Jabatan</Label>
                  <p className="text-sm font-medium">{detailRecord.jabatan}</p>
                </div>
                <div className="space-y-1">
                  <Label className="text-[10px] uppercase tracking-wider text-muted-foreground">Unit Kerja</Label>
                  <p className="text-sm font-medium">{detailRecord.unitKerja}</p>
                </div>
                <div className="space-y-1">
                  <Label className="text-[10px] uppercase tracking-wider text-muted-foreground">Tahun & Periode</Label>
                  <p className="text-sm font-medium">{detailRecord.tahun} - {detailRecord.periode}</p>
                </div>
                <div className="space-y-1">
                  <Label className="text-[10px] uppercase tracking-wider text-muted-foreground">Status Berkas</Label>
                  <div>
                    <Badge variant={detailRecord.status === 'Sudah Diambil' ? 'default' : 'secondary'} className="text-[10px]">
                      {detailRecord.status}
                    </Badge>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4 p-4 bg-primary/5 rounded-lg border border-primary/10">
                <div className="space-y-1">
                  <Label className="text-[10px] uppercase tracking-wider text-muted-foreground">Rating Hasil Kerja</Label>
                  <p className="text-sm font-bold text-primary">{detailRecord.ratingHasilKerja || '-'}</p>
                </div>
                <div className="space-y-1">
                  <Label className="text-[10px] uppercase tracking-wider text-muted-foreground">Rating Hasil Perilaku</Label>
                  <p className="text-sm font-bold text-primary">{detailRecord.ratingHasilPerilaku || '-'}</p>
                </div>
                <div className="space-y-1 col-span-2">
                  <Label className="text-[10px] uppercase tracking-wider text-muted-foreground">Predikat Kinerja</Label>
                  <Badge variant="outline" className="text-sm font-bold border-primary/30 text-primary bg-primary/5">
                    {detailRecord.predikatKinerja || '-'}
                  </Badge>
                </div>
              </div>

              <div className="space-y-2">
                <Label className="text-[10px] uppercase tracking-wider text-muted-foreground">Jenis Dokumen</Label>
                <div className="flex flex-wrap gap-2">
                  {detailRecord.jenisDokumen.map((doc, idx) => (
                    <Badge key={idx} variant="outline" className="text-[10px] bg-primary/5">
                      {doc}
                    </Badge>
                  ))}
                  {detailRecord.jenisDokumen.length === 0 && (
                    <span className="text-xs text-muted-foreground italic">Tidak ada dokumen terpilih</span>
                  )}
                </div>
              </div>

              {detailRecord.status === 'Sudah Diambil' && (
                <div className="p-4 bg-primary/5 rounded-lg border border-primary/10 space-y-3">
                  <div className="flex items-center gap-2 text-primary">
                    <CheckCircle2 className="h-4 w-4" />
                    <span className="text-xs font-bold uppercase tracking-wide">Informasi Pengambilan</span>
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                      <Label className="text-[10px] text-muted-foreground">Nama Pengambil</Label>
                      <p className="text-sm font-semibold">{detailRecord.pengambil}</p>
                    </div>
                    <div className="space-y-1">
                      <Label className="text-[10px] text-muted-foreground">Unit Kerja Pengambil</Label>
                      <p className="text-sm font-semibold">{detailRecord.unitKerjaPengambil}</p>
                    </div>
                    <div className="space-y-1 col-span-2">
                      <Label className="text-[10px] text-muted-foreground">Tanggal Pengambilan</Label>
                      <p className="text-sm font-semibold">
                        {detailRecord.tanggalAmbil ? format(new Date(detailRecord.tanggalAmbil), 'dd MMMM yyyy, HH:mm', { locale: id }) : '-'}
                      </p>
                    </div>
                  </div>
                </div>
              )}

              {detailRecord.fileUrl ? (
                <div className="space-y-3">
                  <Label className="text-[10px] uppercase tracking-wider text-muted-foreground">Berkas Digital</Label>
                  <div className="flex gap-2">
                    <Button 
                      variant="secondary" 
                      size="sm" 
                      className="flex-1 gap-2 bg-primary/10 text-primary hover:bg-primary/20 border-primary/20"
                      onClick={() => {
                        setViewingFileUrl(detailRecord.fileUrl || null);
                        setViewingFileName(detailRecord.fileName || 'berkas-skp');
                        setIsViewFileOpen(true);
                      }}
                    >
                      <Eye className="h-4 w-4" />
                      Lihat Berkas
                    </Button>
                    <Button 
                      variant="outline" 
                      size="sm" 
                      className="flex-1 gap-2"
                      nativeButton={false}
                      render={
                        <a href={detailRecord.fileUrl} download={detailRecord.fileName || 'berkas-skp'}>
                          <Download className="h-4 w-4" />
                          Unduh Berkas
                        </a>
                      }
                    />
                  </div>
                </div>
              ) : (
                <div className="space-y-2">
                  <Label className="text-[10px] uppercase tracking-wider text-muted-foreground">Berkas Digital</Label>
                  <div className="p-3 rounded-lg border border-dashed flex items-center justify-center bg-muted/20">
                    <span className="text-xs text-muted-foreground italic">Belum ada berkas digital yang diunggah</span>
                  </div>
                </div>
              )}
            </div>
          )}
          <DialogFooter>
            <Button 
              variant="outline" 
              className="mr-auto"
              onClick={() => {
                setIsDetailDialogOpen(false);
                setEditingRecord(detailRecord);
                setIsEditDialogOpen(true);
              }}
            >
              Edit Data
            </Button>
            <DialogClose 
              nativeButton={true}
              render={<Button variant="secondary">Tutup</Button>} 
            />
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* View File Dialog */}
      <Dialog open={isViewFileOpen} onOpenChange={setIsViewFileOpen}>
        <DialogContent className="sm:max-w-[800px] h-[80vh] flex flex-col p-0 overflow-hidden">
          <DialogHeader className="p-6 border-b">
            <DialogTitle className="flex items-center gap-2">
              <FileText className="h-5 w-5 text-primary" />
              Pratinjau Berkas: {viewingFileName}
            </DialogTitle>
          </DialogHeader>
          <div className="flex-1 bg-muted/30 overflow-auto p-4 flex items-center justify-center">
            {viewingFileUrl ? (
              viewingFileName?.toLowerCase().endsWith('.pdf') ? (
                <div className="w-full h-full flex flex-col gap-4">
                  <object 
                    data={viewingFileUrl} 
                    type="application/pdf"
                    className="w-full h-full rounded border shadow-sm"
                  >
                    <div className="flex flex-col items-center justify-center h-full p-8 text-center bg-background rounded border">
                      <FileText className="h-12 w-12 text-muted-foreground mb-4" />
                      <p className="text-sm font-medium mb-2">Pratinjau PDF tidak dapat ditampilkan secara otomatis</p>
                      <p className="text-xs text-muted-foreground mb-4">Browser Anda mungkin memblokir pratinjau berkas PDF di dalam jendela ini.</p>
                      <Button 
                        variant="outline"
                        nativeButton={false}
                        render={
                          <a href={viewingFileUrl} target="_blank" rel="noopener noreferrer">
                            Buka PDF di Tab Baru
                          </a>
                        }
                      />
                    </div>
                  </object>
                </div>
              ) : (
                <img 
                  src={viewingFileUrl} 
                  alt="Preview" 
                  className="max-w-full max-h-full object-contain rounded shadow-sm"
                  referrerPolicy="no-referrer"
                />
              )
            ) : (
              <p className="text-muted-foreground">Gagal memuat berkas.</p>
            )}
          </div>
          <DialogFooter className="p-4 border-t bg-white">
            <DialogClose 
              nativeButton={true}
              render={<Button variant="outline" />}
            >
              Tutup
            </DialogClose>
            <Button 
              nativeButton={false}
              render={<a href={viewingFileUrl || '#'} download={viewingFileName || 'berkas-skp'} />}
            >
              <Download className="h-4 w-4 mr-2" />
              Unduh Berkas
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <footer className="container mx-auto p-8 text-center text-xs text-muted-foreground">
        <p>&copy; {new Date().getFullYear()} Monitoring SKP Pegawai. Sistem Monitoring Berkas Kepegawaian.</p>
      </footer>
    </div>
  );
}
