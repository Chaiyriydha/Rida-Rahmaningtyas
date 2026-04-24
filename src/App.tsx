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
  ArrowRight,
  User,
  FileText,
  Eye,
  UploadCloud,
  X,
  UserCheck,
  UserCog,
  LogOut,
  LogIn,
  AlertTriangle,
  Pencil
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { format } from 'date-fns';
import { id } from 'date-fns/locale';
import * as XLSX from 'xlsx';
import { 
  collection, 
  addDoc, 
  updateDoc, 
  deleteDoc, 
  doc, 
  onSnapshot, 
  query, 
  orderBy,
  serverTimestamp,
  Timestamp
} from 'firebase/firestore';
import { 
  signInWithPopup, 
  GoogleAuthProvider, 
  onAuthStateChanged, 
  signOut,
  User as FirebaseUser
} from 'firebase/auth';

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

import { SKPRecord, SKPStatus, SKPPeriode, PLTRecord, PLHRecord, PLTStatus, WorkTeam, TeamMember } from './types';

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId: string | undefined;
    email: string | null | undefined;
    emailVerified: boolean | undefined;
    isAnonymous: boolean | undefined;
    tenantId: string | null | undefined;
    providerInfo: {
      providerId: string;
      displayName: string | null;
      email: string | null;
      photoUrl: string | null;
    }[];
  }
}
import { 
  UNIT_KERJA_LIST, 
  PERIODE_LIST, 
  JENIS_DOKUMEN_LIST,
  RATING_HASIL_KERJA_LIST,
  RATING_PERILAKU_LIST,
  PREDIKAT_KINERJA_LIST,
  STATUS_PEGAWAI_LIST
} from './constants';
import { cn } from '@/lib/utils';
import { db, auth } from './lib/firebase';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  Legend, 
  ResponsiveContainer 
} from 'recharts';

export default function App() {
  const [activeTab, setActiveTab] = useState<'tim-kerja' | 'skp' | 'plt-plh'>('tim-kerja');
  const [user, setUser] = useState<FirebaseUser | null>(null);
  const [isAuthLoading, setIsAuthLoading] = useState(true);
  const [records, setRecords] = useState<SKPRecord[]>([]);
  const [pltRecords, setPltRecords] = useState<PLTRecord[]>([]);
  const [plhRecords, setPlhRecords] = useState<PLHRecord[]>([]);
  const [workTeams, setWorkTeams] = useState<WorkTeam[]>([]);
  const [searchQuery, setSearchQuery] = useState('');

  const groupedTeams = useMemo(() => {
    const groups: Record<string, { topLevel: WorkTeam[], subTeams: Record<string, WorkTeam[]> }> = {};
    workTeams.forEach(team => {
      if (!groups[team.unitKerja]) {
        groups[team.unitKerja] = { topLevel: [], subTeams: {} };
      }
      
      if (!team.parentId) {
        groups[team.unitKerja].topLevel.push(team);
      } else {
        if (!groups[team.unitKerja].subTeams[team.parentId]) {
          groups[team.unitKerja].subTeams[team.parentId] = [];
        }
        groups[team.unitKerja].subTeams[team.parentId].push(team);
      }
    });
    return groups;
  }, [workTeams]);
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [statusFilterPlt, setStatusFilterPlt] = useState<string>('all');
  const [statusFilterPlh, setStatusFilterPlh] = useState<string>('all');
  const [yearFilterPltPlh, setYearFilterPltPlh] = useState<string>('all');
  const [isAddDialogOpen, setIsAddDialogOpen] = useState(false);
  const [isAddPltPlhDialogOpen, setIsAddPltPlhDialogOpen] = useState(false);
  const [isTakeDialogOpen, setIsTakeDialogOpen] = useState(false);
  const [isDetailDialogOpen, setIsDetailDialogOpen] = useState(false);
  const [isEditDialogOpen, setIsEditDialogOpen] = useState(false);
  const [isViewFileOpen, setIsViewFileOpen] = useState(false);
  const [selectedRecord, setSelectedRecord] = useState<SKPRecord | null>(null);
  const [detailRecord, setDetailRecord] = useState<SKPRecord | null>(null);
  const [editingRecord, setEditingRecord] = useState<SKPRecord | null>(null);
  const [editingPltRecord, setEditingPltRecord] = useState<PLTRecord | null>(null);
  const [editingPlhRecord, setEditingPlhRecord] = useState<PLHRecord | null>(null);
  const [pltPlhType, setPltPlhType] = useState<'PLT' | 'PLH'>('PLT');
  const [isEditPltPlhDialogOpen, setIsEditPltPlhDialogOpen] = useState(false);
  const [isAddTeamDialogOpen, setIsAddTeamDialogOpen] = useState(false);
  const [isEditTeamDialogOpen, setIsEditTeamDialogOpen] = useState(false);
  const [editingTeam, setEditingTeam] = useState<WorkTeam | null>(null);
  const [jenisAddTeam, setJenisAddTeam] = useState<'Tim Kerja' | 'Bagian'>('Tim Kerja');
  const [jenisEditTeam, setJenisEditTeam] = useState<'Tim Kerja' | 'Bagian'>('Tim Kerja');
  const [isAddMemberDialogOpen, setIsAddMemberDialogOpen] = useState(false);
  const [isEditMemberDialogOpen, setIsEditMemberDialogOpen] = useState(false);
  const [selectedTeamForMember, setSelectedTeamForMember] = useState<string | null>(null);
  const [editingMemberInfo, setEditingMemberInfo] = useState<{ teamId: string, member: TeamMember, index: number } | null>(null);
  const [pendingSubTeams, setPendingSubTeams] = useState<{ namaTim: string, ketua: TeamMember, anggota: TeamMember[] }[]>([]);
  const [isDeleteConfirmOpen, setIsDeleteConfirmOpen] = useState(false);
  const [deletingInfo, setDeletingInfo] = useState<{ id: string, type: 'SKP' | 'PLT' | 'PLH' | 'TIM' } | null>(null);
  const [viewingFileUrl, setViewingFileUrl] = useState<string | null>(null);
  const [viewingFileName, setViewingFileName] = useState<string | null>(null);
  const [uploadPreview, setUploadPreview] = useState<string | null>(null);
  const [uploadFileName, setUploadFileName] = useState<string | null>(null);

  // Auth Effect
  React.useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      setIsAuthLoading(false);
    });
    return () => unsubscribe();
  }, []);

  // Firestore Real-time Sync
  React.useEffect(() => {
    if (!user) {
      setRecords([]);
      setPltRecords([]);
      setPlhRecords([]);
      return;
    }

    const qSkp = query(collection(db, 'skp_records'), orderBy('createdAt', 'desc'));
    const unsubSkp = onSnapshot(qSkp, (snapshot) => {
      const docs = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as SKPRecord[];
      setRecords(docs);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'skp_records');
    });

    const qPlt = query(collection(db, 'plt_records'), orderBy('status', 'asc'));
    const unsubPlt = onSnapshot(qPlt, (snapshot) => {
      const docs = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as PLTRecord[];
      setPltRecords(docs);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'plt_records');
    });

    const qPlh = query(collection(db, 'plh_records'), orderBy('status', 'asc'));
    const unsubPlh = onSnapshot(qPlh, (snapshot) => {
      const docs = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as PLHRecord[];
      setPlhRecords(docs);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'plh_records');
    });

    const qWorkTeams = query(collection(db, 'work_teams'), orderBy('namaTim', 'asc'));
    const unsubWorkTeams = onSnapshot(qWorkTeams, (snapshot) => {
      const docs = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as WorkTeam[];
      setWorkTeams(docs);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'work_teams');
    });

    return () => {
      unsubSkp();
      unsubPlt();
      unsubPlh();
      unsubWorkTeams();
    };
  }, [user]);

  const handleLogin = async () => {
    const provider = new GoogleAuthProvider();
    try {
      await signInWithPopup(auth, provider);
      toast.success('Berhasil masuk');
    } catch (error) {
      console.error("Login Error:", error);
      toast.error('Gagal masuk dengan Google');
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
      toast.success('Berhasil keluar');
    } catch (error) {
      console.error("Logout Error:", error);
    }
  };

  // Auto-update status for expired PLT/PLH records
  React.useEffect(() => {
    if (!user) return;

    const checkExpiredRecords = async () => {
      const today = new Date();
      today.setHours(0, 0, 0, 0);

      // Check PLH records
      if (plhRecords.length > 0) {
        for (const record of plhRecords) {
          if (record.status === 'Aktif' && record.tglSelesai) {
            const endDate = new Date(record.tglSelesai);
            endDate.setHours(0, 0, 0, 0);

            if (endDate < today) {
              try {
                await updateDoc(doc(db, 'plh_records', record.id), {
                  status: 'Selesai',
                  updatedAt: serverTimestamp()
                });
              } catch (error) {
                console.error(`Failed to auto-update PLH record ${record.id}:`, error);
              }
            }
          }
        }
      }

      // Check PLT records (if they have tglSelesai)
      if (pltRecords.length > 0) {
        for (const record of pltRecords) {
          if (record.status === 'Aktif' && record.tglSelesai) {
            const endDate = new Date(record.tglSelesai);
            endDate.setHours(0, 0, 0, 0);

            if (endDate < today) {
              try {
                await updateDoc(doc(db, 'plt_records', record.id), {
                  status: 'Selesai',
                  updatedAt: serverTimestamp()
                });
              } catch (error) {
                console.error(`Failed to auto-update PLT record ${record.id}:`, error);
              }
            }
          }
        }
      }
    };

    checkExpiredRecords();
  }, [user, plhRecords, pltRecords]);

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
    return pltRecords.filter(record => {
      const matchesSearch = 
        record.namaPegawai.toLowerCase().includes(searchQuery.toLowerCase()) ||
        record.nip.includes(searchQuery) ||
        record.jabatanPlt.toLowerCase().includes(searchQuery.toLowerCase()) ||
        record.unitKerja.toLowerCase().includes(searchQuery.toLowerCase());
      
      const recordYear = new Date(record.tglMulai).getFullYear().toString();
      const matchesYear = yearFilterPltPlh === 'all' || recordYear === yearFilterPltPlh;
      const matchesStatus = statusFilterPlt === 'all' || record.status === statusFilterPlt;
      
      return matchesSearch && matchesYear && matchesStatus;
    });
  }, [pltRecords, searchQuery, yearFilterPltPlh, statusFilterPlt]);

  const filteredPlhRecords = useMemo(() => {
    return plhRecords.filter(record => {
      const matchesSearch = 
        record.namaPegawai.toLowerCase().includes(searchQuery.toLowerCase()) ||
        record.nip.includes(searchQuery) ||
        record.jabatanPlh.toLowerCase().includes(searchQuery.toLowerCase()) ||
        record.unitKerja.toLowerCase().includes(searchQuery.toLowerCase());
      
      const recordYear = new Date(record.tglMulai).getFullYear().toString();
      const matchesYear = yearFilterPltPlh === 'all' || recordYear === yearFilterPltPlh;
      const matchesStatus = statusFilterPlh === 'all' || record.status === statusFilterPlh;
      
      return matchesSearch && matchesYear && matchesStatus;
    });
  }, [plhRecords, searchQuery, yearFilterPltPlh, statusFilterPlh]);

  const chartDataPltPlh = useMemo(() => {
    const years = new Set<number>();
    const currentYear = new Date().getFullYear();
    
    // Get unique years from both records
    pltRecords.forEach(r => years.add(new Date(r.tglMulai).getFullYear()));
    plhRecords.forEach(r => years.add(new Date(r.tglMulai).getFullYear()));
    
    // Ensure at least the last 3 years are shown if no data
    if (years.size === 0) {
      years.add(currentYear);
      years.add(currentYear - 1);
      years.add(currentYear - 2);
    }

    const sortedYears = Array.from(years).sort((a, b) => a - b);
    
    return sortedYears.map(year => {
      const pltCount = pltRecords.filter(r => new Date(r.tglMulai).getFullYear() === year).length;
      const plhCount = plhRecords.filter(r => new Date(r.tglMulai).getFullYear() === year).length;
      return {
        year: year.toString(),
        PLT: pltCount,
        PLH: plhCount
      };
    });
  }, [pltRecords, plhRecords]);

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

  const handleAddRecord = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!user) return;

    const formData = new FormData(e.currentTarget);
    const selectedJenisDokumen = JENIS_DOKUMEN_LIST.filter(item => formData.get(`jenisDokumen-${item}`) === 'on');
    
    const fileInput = e.currentTarget.querySelector('input[type="file"]') as HTMLInputElement;
    const file = fileInput?.files?.[0];
    
    let fileUrl = '';
    let fileName = '';
    
    // Note: In a real app, we would upload to Firebase Storage here.
    // For now, we'll use object URLs as placeholders or just store the metadata.
    if (file) {
      fileUrl = URL.createObjectURL(file);
      fileName = file.name;
    }

    try {
      const newRecord = {
        nip: formData.get('nip') as string,
        namaPegawai: formData.get('namaPegawai') as string,
        jabatan: formData.get('jabatan') as string,
        unitKerja: formData.get('unitKerja') as string,
        tahun: parseInt(formData.get('tahun') as string),
        periode: formData.get('periode') as SKPPeriode,
        jenisDokumen: selectedJenisDokumen,
        status: 'Belum Diambil' as SKPStatus,
        ratingHasilKerja: formData.get('ratingHasilKerja') as string,
        ratingHasilPerilaku: formData.get('ratingHasilPerilaku') as string,
        predikatKinerja: formData.get('predikatKinerja') as string,
        fileUrl,
        fileName,
        authorUid: user.uid,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp(),
      };

      await addDoc(collection(db, 'skp_records'), newRecord);
      setIsAddDialogOpen(false);
      setUploadPreview(null);
      setUploadFileName(null);
      toast.success('Data SKP berhasil disimpan ke database');
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, 'skp_records');
    }
  };

  const handleAddPltPlhRecord = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!user) return;

    const formData = new FormData(e.currentTarget);
    const type = formData.get('type') as 'PLT' | 'PLH';
    const collectionName = type === 'PLT' ? 'plt_records' : 'plh_records';

    try {
      const newRecord = {
        nip: formData.get('nip') as string,
        namaPegawai: formData.get('namaPegawai') as string,
        jabatanAsli: formData.get('jabatanAsli') as string,
        [type === 'PLT' ? 'jabatanPlt' : 'jabatanPlh']: formData.get('jabatanTugas') as string,
        unitKerja: formData.get('unitKerja') as string,
        unitKerjaTugas: formData.get('unitKerjaTugas') as string,
        noSk: formData.get('noSk') as string,
        tglMulai: formData.get('tglMulai') as string,
        ...(type === 'PLH' && { tglSelesai: formData.get('tglSelesai') as string }),
        status: 'Aktif' as PLTStatus,
        keterangan: formData.get('keterangan') as string,
        pegawaiDigantikan: formData.get('pegawaiDigantikan') as string,
        authorUid: user.uid,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp(),
      };

      await addDoc(collection(db, collectionName), newRecord);
      setIsAddPltPlhDialogOpen(false);
      toast.success(`Data ${type} berhasil disimpan`);
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, collectionName);
    }
  };

  const handleUpdateRecord = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!editingRecord || !user) return;

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

    try {
      const updatedData = {
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
        updatedAt: serverTimestamp(),
      };

      await updateDoc(doc(db, 'skp_records', editingRecord.id), updatedData);
      setIsEditDialogOpen(false);
      setEditingRecord(null);
      setUploadPreview(null);
      setUploadFileName(null);
      toast.success('Data SKP berhasil diperbarui');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `skp_records/${editingRecord.id}`);
    }
  };

  const handleUpdatePltPlhRecord = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!user) return;
    
    const editingData = editingPltRecord || editingPlhRecord;
    if (!editingData) return;

    const formData = new FormData(e.currentTarget);
    const type = formData.get('type') as 'PLT' | 'PLH';
    const collectionName = type === 'PLT' ? 'plt_records' : 'plh_records';

    try {
      const updatedData = {
        nip: formData.get('nip') as string,
        namaPegawai: formData.get('namaPegawai') as string,
        jabatanAsli: formData.get('jabatanAsli') as string,
        [type === 'PLT' ? 'jabatanPlt' : 'jabatanPlh']: formData.get('jabatanTugas') as string,
        unitKerja: formData.get('unitKerja') as string,
        unitKerjaTugas: formData.get('unitKerjaTugas') as string,
        noSk: formData.get('noSk') as string,
        tglMulai: formData.get('tglMulai') as string,
        ...(type === 'PLH' && { tglSelesai: formData.get('tglSelesai') as string }),
        status: formData.get('status') as PLTStatus,
        keterangan: formData.get('keterangan') as string,
        pegawaiDigantikan: formData.get('pegawaiDigantikan') as string,
        updatedAt: serverTimestamp(),
      };

      await updateDoc(doc(db, collectionName, editingData.id), updatedData);
      setIsEditPltPlhDialogOpen(false);
      setEditingPltRecord(null);
      setEditingPlhRecord(null);
      toast.success(`Data ${type} berhasil diperbarui`);
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `${collectionName}/${editingData.id}`);
    }
  };

  const handleAddTeam = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!user) return;

    const formData = new FormData(e.currentTarget);
    const namaTim = formData.get('namaTim') as string;
    const unitKerja = formData.get('unitKerja') as string;
    const tahun = parseInt(formData.get('tahun') as string);
    const ketuaNama = formData.get('ketuaNama') as string;
    const ketuaJabatan = formData.get('ketuaJabatan') as string;
    const ketuaStatus = formData.get('ketuaStatus') as string;
    const jenis = formData.get('jenis') as 'Tim Kerja' | 'Bagian';
    const parentId = formData.get('parentId') as string;

    try {
      const newTeam = {
        namaTim,
        unitKerja,
        tahun,
        status: 'Aktif',
        jenis,
        parentId: parentId || null,
        ketua: {
          nama: ketuaNama,
          jabatan: ketuaJabatan,
          status: ketuaStatus
        },
        anggota: [], // Initial empty members, can be added via edit
        authorUid: user.uid,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp(),
      };

      const docRef = await addDoc(collection(db, 'work_teams'), newTeam);
      
      // If Bagian, also add sub-teams
      if (jenis === 'Bagian' && pendingSubTeams.length > 0) {
        for (const sub of pendingSubTeams) {
          await addDoc(collection(db, 'work_teams'), {
            namaTim: sub.namaTim,
            unitKerja,
            tahun,
            status: 'Aktif',
            jenis: 'Tim Kerja',
            parentId: docRef.id,
            ketua: sub.ketua,
            anggota: sub.anggota,
            authorUid: user.uid,
            createdAt: serverTimestamp(),
            updatedAt: serverTimestamp(),
          });
        }
      }

      setPendingSubTeams([]);
      setIsAddTeamDialogOpen(false);
      toast.success(jenis === 'Bagian' ? 'Bagian dan Tim Kerja berhasil ditambahkan' : 'Tim Kerja berhasil ditambahkan');
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, 'work_teams');
    }
  };

  const handleUpdateTeam = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!user || !editingTeam) return;

    const formData = new FormData(e.currentTarget);
    const namaTim = formData.get('namaTim') as string;
    const unitKerja = formData.get('unitKerja') as string;
    const tahun = parseInt(formData.get('tahun') as string);
    const status = formData.get('status') as 'Aktif' | 'Non-Aktif';
    const ketuaNama = formData.get('ketuaNama') as string;
    const ketuaJabatan = formData.get('ketuaJabatan') as string;
    const ketuaStatus = formData.get('ketuaStatus') as string;
    const jenis = formData.get('jenis') as 'Tim Kerja' | 'Bagian';
    const parentId = formData.get('parentId') as string;

    try {
      const updatedData = {
        namaTim,
        unitKerja,
        tahun,
        status,
        jenis,
        parentId: parentId || null,
        ketua: {
          nama: ketuaNama,
          jabatan: ketuaJabatan,
          status: ketuaStatus
        },
        updatedAt: serverTimestamp(),
      };

      await updateDoc(doc(db, 'work_teams', editingTeam.id), updatedData);
      setIsEditTeamDialogOpen(false);
      setEditingTeam(null);
      toast.success('Tim Kerja berhasil diperbarui');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `work_teams/${editingTeam.id}`);
    }
  };

  const handleAddMember = async (teamId: string, member: TeamMember) => {
    if (!user) return;
    const team = workTeams.find(t => t.id === teamId);
    if (!team) return;

    try {
      const updatedAnggota = [...team.anggota, member];
      await updateDoc(doc(db, 'work_teams', teamId), {
        anggota: updatedAnggota,
        updatedAt: serverTimestamp()
      });
      toast.success('Anggota berhasil ditambahkan');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `work_teams/${teamId}`);
    }
  };

  const handleRemoveMember = async (teamId: string, memberNama: string) => {
    if (!user) return;
    const team = workTeams.find(t => t.id === teamId);
    if (!team) return;

    try {
      const updatedAnggota = team.anggota.filter(m => m.nama !== memberNama);
      await updateDoc(doc(db, 'work_teams', teamId), {
        anggota: updatedAnggota,
        updatedAt: serverTimestamp()
      });
      toast.success('Anggota berhasil dihapus');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `work_teams/${teamId}`);
    }
  };

  const handleUpdateMember = async (teamId: string, memberIndex: number, updatedMember: TeamMember) => {
    if (!user) return;
    const team = workTeams.find(t => t.id === teamId);
    if (!team) return;

    try {
      const updatedAnggota = [...team.anggota];
      updatedAnggota[memberIndex] = updatedMember;
      await updateDoc(doc(db, 'work_teams', teamId), {
        anggota: updatedAnggota,
        updatedAt: serverTimestamp()
      });
      toast.success('Informasi anggota berhasil diperbarui');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `work_teams/${teamId}`);
    }
  };

  const addPendingSubTeam = () => {
    setPendingSubTeams([...pendingSubTeams, { 
      namaTim: '', 
      ketua: { nama: '', jabatan: '', status: 'PNS' }, 
      anggota: [] 
    }]);
  };

  const removePendingSubTeam = (index: number) => {
    setPendingSubTeams(pendingSubTeams.filter((_, i) => i !== index));
  };

  const updatePendingSubTeam = (index: number, field: string, value: any) => {
    const newSubTeams = [...pendingSubTeams];
    if (field.startsWith('ketua.')) {
      const subField = field.split('.')[1] as keyof TeamMember;
      newSubTeams[index].ketua[subField] = value;
    } else {
      (newSubTeams[index] as any)[field] = value;
    }
    setPendingSubTeams(newSubTeams);
  };

  const addMemberToPendingSubTeam = (subTeamIndex: number) => {
    const newSubTeams = [...pendingSubTeams];
    newSubTeams[subTeamIndex].anggota.push({ nama: '', jabatan: '', status: 'PNS' });
    setPendingSubTeams(newSubTeams);
  };

  const removeMemberFromPendingSubTeam = (subTeamIndex: number, memberIndex: number) => {
    const newSubTeams = [...pendingSubTeams];
    newSubTeams[subTeamIndex].anggota = newSubTeams[subTeamIndex].anggota.filter((_, i) => i !== memberIndex);
    setPendingSubTeams(newSubTeams);
  };

  const updateMemberInPendingSubTeam = (subTeamIndex: number, memberIndex: number, field: keyof TeamMember, value: string) => {
    const newSubTeams = [...pendingSubTeams];
    newSubTeams[subTeamIndex].anggota[memberIndex][field] = value;
    setPendingSubTeams(newSubTeams);
  };

  const handleTakeSKP = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!selectedRecord || !user) return;

    const formData = new FormData(e.currentTarget);
    const tanggalAmbilInput = formData.get('tanggalAmbil') as string;
    
    try {
      const updatedData = {
        status: 'Sudah Diambil' as SKPStatus,
        pengambil: formData.get('pengambil') as string,
        unitKerjaPengambil: formData.get('unitKerjaPengambil') as string,
        tanggalAmbil: new Date(tanggalAmbilInput).toISOString(),
        updatedAt: serverTimestamp(),
      };

      await updateDoc(doc(db, 'skp_records', selectedRecord.id), updatedData);
      setIsTakeDialogOpen(false);
      setSelectedRecord(null);
      toast.success('Status berkas berhasil diperbarui');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `skp_records/${selectedRecord.id}`);
    }
  };

  const handleDeleteRecord = (id: string) => {
    setDeletingInfo({ id, type: 'SKP' });
    setIsDeleteConfirmOpen(true);
  };

  const handleDeletePltPlhRecord = (id: string, type: 'PLT' | 'PLH') => {
    setDeletingInfo({ id, type });
    setIsDeleteConfirmOpen(true);
  };

  const handleFirestoreError = (error: unknown, operationType: OperationType, path: string | null) => {
    const errInfo: FirestoreErrorInfo = {
      error: error instanceof Error ? error.message : String(error),
      authInfo: {
        userId: auth.currentUser?.uid,
        email: auth.currentUser?.email,
        emailVerified: auth.currentUser?.emailVerified,
        isAnonymous: auth.currentUser?.isAnonymous,
        tenantId: auth.currentUser?.tenantId,
        providerInfo: auth.currentUser?.providerData.map(provider => ({
          providerId: provider.providerId,
          displayName: provider.displayName,
          email: provider.email,
          photoUrl: provider.photoURL
        })) || []
      },
      operationType,
      path
    }
    console.error('Firestore Error: ', JSON.stringify(errInfo));
    // Check if it's a permission error
    if (errInfo.error.toLowerCase().includes('permission') || errInfo.error.toLowerCase().includes('insufficient')) {
      toast.error("Akses ditolak: Anda tidak memiliki izin untuk melakukan tindakan ini.");
    }
    throw new Error(JSON.stringify(errInfo));
  };

  const confirmDelete = async () => {
    if (!deletingInfo) return;
    const { id, type } = deletingInfo;
    const collectionName = type === 'SKP' ? 'skp_records' : (type === 'PLT' ? 'plt_records' : (type === 'PLH' ? 'plh_records' : 'work_teams'));
    
    try {
      await deleteDoc(doc(db, collectionName, id));
      toast.success(`Data ${type === 'SKP' ? '' : type} berhasil dihapus`);
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, `${collectionName}/${id}`);
    } finally {
      setIsDeleteConfirmOpen(false);
      setDeletingInfo(null);
    }
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

  const handleExportPltPlh = () => {
    if (pltRecords.length === 0 && plhRecords.length === 0) {
      toast.error('Tidak ada data PLT/PLH untuk diekspor');
      return;
    }

    const workbook = XLSX.utils.book_new();

    if (pltRecords.length > 0) {
      const pltData = pltRecords.map(record => ({
        'Nama Pegawai': record.namaPegawai,
        'NIP': record.nip,
        'Jabatan Asli': record.jabatanAsli,
        'Jabatan PLT': record.jabatanPlt,
        'Unit Kerja Asal': record.unitKerja,
        'Unit Kerja Penugasan': record.unitKerjaTugas,
        'Nomor SK': record.noSk,
        'Alasan/Keterangan': record.keterangan || '-',
        'Tanggal Mulai': format(new Date(record.tglMulai), 'dd MMMM yyyy', { locale: id }),
        'Tanggal Selesai': record.tglSelesai ? format(new Date(record.tglSelesai), 'dd MMMM yyyy', { locale: id }) : 'Selesai',
        'Status': record.status
      }));
      const pltSheet = XLSX.utils.json_to_sheet(pltData);
      XLSX.utils.book_append_sheet(workbook, pltSheet, 'Data PLT');
      
      // Auto-size columns for PLT
      const pltWidths = Object.keys(pltData[0]).map(key => ({
        wch: Math.max(key.length, ...pltData.map(row => String(row[key as keyof typeof row]).length)) + 2
      }));
      pltSheet['!cols'] = pltWidths;
    }

    if (plhRecords.length > 0) {
      const plhData = plhRecords.map(record => ({
        'Nama Pegawai': record.namaPegawai,
        'NIP': record.nip,
        'Jabatan Asli': record.jabatanAsli,
        'Jabatan PLH': record.jabatanPlh,
        'Unit Kerja Asal': record.unitKerja,
        'Unit Kerja Penugasan': record.unitKerjaTugas,
        'Nomor SK': record.noSk,
        'Alasan/Keterangan': record.keterangan || '-',
        'Tanggal Mulai': format(new Date(record.tglMulai), 'dd MMMM yyyy', { locale: id }),
        'Tanggal Selesai': format(new Date(record.tglSelesai), 'dd MMMM yyyy', { locale: id }),
        'Status': record.status
      }));
      const plhSheet = XLSX.utils.json_to_sheet(plhData);
      XLSX.utils.book_append_sheet(workbook, plhSheet, 'Data PLH');

      // Auto-size columns for PLH
      const plhWidths = Object.keys(plhData[0]).map(key => ({
        wch: Math.max(key.length, ...plhData.map(row => String(row[key as keyof typeof row]).length)) + 2
      }));
      plhSheet['!cols'] = plhWidths;
    }

    XLSX.writeFile(workbook, `Data_PLT_PLH_Export_${format(new Date(), 'yyyyMMdd_HHmm')}.xlsx`);
    toast.success('Data PLT & PLH berhasil diekspor ke Excel');
  };

  const renderTeamCard = (team: WorkTeam, isSubTeam = false) => (
    <Card key={team.id} className={`overflow-hidden border-none shadow-sm hover:shadow-md transition-all group ${isSubTeam ? 'border-l-4 border-l-primary/20 bg-primary/5' : ''}`}>
      <CardHeader className={`${isSubTeam ? 'bg-primary/10' : 'bg-primary/5'} pb-4`}>
        <div className="flex justify-between items-start">
          <div className="space-y-1">
            <div className="flex items-center gap-2">
              <CardTitle className="text-base font-bold text-primary group-hover:text-primary/80 transition-colors">
                {team.namaTim}
              </CardTitle>
              {team.jenis === 'Bagian' && (
                <Badge variant="outline" className="text-[8px] h-4 px-1 bg-primary/10 text-primary border-primary/20">
                  Bagian
                </Badge>
              )}
            </div>
            <CardDescription className="text-[10px] font-medium uppercase tracking-wider">
              Tahun {team.tahun}
            </CardDescription>
          </div>
          <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
            <Button 
              variant="ghost" 
              size="icon" 
              className="h-7 w-7 text-muted-foreground hover:text-primary"
              onClick={() => {
                setEditingTeam(team);
                setJenisEditTeam(team.jenis || 'Tim Kerja');
                setIsEditTeamDialogOpen(true);
              }}
            >
              <UserCog className="h-3.5 w-3.5" />
            </Button>
            <Button 
              variant="ghost" 
              size="icon" 
              className="h-7 w-7 text-muted-foreground hover:text-destructive"
              onClick={() => {
                setDeletingInfo({ id: team.id, type: 'TIM' });
                setIsDeleteConfirmOpen(true);
              }}
            >
              <X className="h-3.5 w-3.5" />
            </Button>
          </div>
        </div>
      </CardHeader>
      <CardContent className="pt-4 space-y-4">
        {/* Ketua Section */}
        <div className="space-y-2">
          <div className="flex items-center gap-2 text-[10px] font-bold text-muted-foreground uppercase tracking-widest">
            <UserCheck className="h-3 w-3" />
            {team.jenis === 'Bagian' ? 'Kepala Bagian' : 'Ketua Tim'}
          </div>
          <div className="flex flex-col bg-muted/30 p-2 rounded-lg border border-muted-foreground/5">
            <div className="flex justify-between items-start">
              <span className="text-sm font-bold">{team.ketua.nama}</span>
              <Badge variant="secondary" className="text-[8px] h-3.5 px-1 py-0 bg-primary/10 text-primary border-none">
                {team.ketua.status}
              </Badge>
            </div>
            <span className="text-[10px] text-muted-foreground font-medium italic">{team.ketua.jabatan}</span>
          </div>
        </div>

        {/* Anggota Section */}
        <div className="space-y-2">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2 text-[10px] font-bold text-muted-foreground uppercase tracking-widest">
              <Users className="h-3 w-3" />
              Anggota ({team.anggota.length})
            </div>
            <Button 
              variant="ghost" 
              size="sm" 
              className="h-6 px-2 text-[10px] font-bold text-primary hover:bg-primary/10"
              onClick={() => {
                setSelectedTeamForMember(team.id);
                setIsAddMemberDialogOpen(true);
              }}
            >
              <Plus className="h-3 w-3 mr-1" />
              Tambah
            </Button>
          </div>
          
          <div className="space-y-1.5 max-h-[200px] overflow-y-auto pr-1 custom-scrollbar">
            {team.anggota.length > 0 ? (
              team.anggota.map((member) => (
                <div key={member.nama} className="flex items-center justify-between group/member bg-white p-2 rounded-md border border-muted shadow-sm hover:border-primary/20 transition-colors">
                  <div className="flex flex-col min-w-0">
                    <span className="text-xs font-bold truncate">{member.nama}</span>
                    <span className="text-[10px] text-muted-foreground font-medium italic leading-tight">{member.jabatan}</span>
                    <div className="mt-1">
                      <Badge variant="outline" className="text-[8px] h-3.5 px-1 py-0 border-muted-foreground/20 text-muted-foreground">
                        {member.status}
                      </Badge>
                    </div>
                  </div>
                  <div className="flex items-center gap-1 opacity-0 group-hover/member:opacity-100 transition-opacity">
                    <Button 
                      variant="ghost" 
                      size="icon" 
                      className="h-6 w-6 text-muted-foreground hover:text-primary"
                      onClick={() => {
                        setEditingMemberInfo({ 
                          teamId: team.id, 
                          member: member, 
                          index: team.anggota.indexOf(member) 
                        });
                        setIsEditMemberDialogOpen(true);
                      }}
                    >
                      <Pencil className="h-3 w-3" />
                    </Button>
                    <Button 
                      variant="ghost" 
                      size="icon" 
                      className="h-6 w-6 text-muted-foreground hover:text-destructive"
                      onClick={() => handleRemoveMember(team.id, member.nama)}
                    >
                      <X className="h-3 w-3" />
                    </Button>
                  </div>
                </div>
              ))
            ) : (
              <div className="text-center py-4 border border-dashed rounded-lg">
                <p className="text-[10px] text-muted-foreground italic">Belum ada anggota</p>
              </div>
            )}
          </div>
        </div>
      </CardContent>
    </Card>
  );

  return (
    <div className="flex min-h-screen bg-[#F8F9FA] text-[#1A1A1A] font-sans selection:bg-primary/20">
      <Toaster position="top-right" />
      
      {/* Sidebar */}
      <aside className="hidden md:flex w-64 flex-col border-r bg-white sticky top-0 h-screen">
        <div className="p-6 flex flex-col gap-8">
          <div className="flex items-center gap-3">
            <div className="flex h-14 w-auto items-center justify-center border-r pr-3 border-muted-foreground/20">
              <span className="text-4xl font-black tracking-tighter text-[#1A4A9A]">BSKJI</span>
            </div>
            <div className="flex flex-col justify-center">
              <p className="text-[8px] font-bold text-[#1A4A9A] leading-tight uppercase">Badan</p>
              <p className="text-[8px] font-bold text-[#1A4A9A] leading-tight uppercase">Standardisasi dan</p>
              <p className="text-[8px] font-bold text-[#1A4A9A] leading-tight uppercase">Kebijakan</p>
              <p className="text-[8px] font-bold text-[#1A4A9A] leading-tight uppercase">Jasa Industri</p>
            </div>
          </div>

          <nav className="flex flex-col gap-2">
            <Button 
              variant={activeTab === 'tim-kerja' ? 'default' : 'ghost'} 
              className={cn(
                "justify-start gap-3 h-10 px-4 transition-all", 
                activeTab === 'tim-kerja' ? "shadow-sm" : "text-muted-foreground"
              )}
              onClick={() => setActiveTab('tim-kerja')}
            >
              <Users className="h-4 w-4" />
              Tim Kerja
            </Button>
            <Button 
              variant={activeTab === 'skp' ? 'default' : 'ghost'} 
              className={cn(
                "justify-start gap-3 h-10 px-4 transition-all", 
                activeTab === 'skp' ? "shadow-sm" : "text-muted-foreground"
              )}
              onClick={() => setActiveTab('skp')}
            >
              <FileText className="h-4 w-4" />
              SKP Pegawai
            </Button>
            <Button 
              variant={activeTab === 'plt-plh' ? 'default' : 'ghost'} 
              className={cn(
                "justify-start gap-3 h-10 px-4 transition-all", 
                activeTab === 'plt-plh' ? "shadow-sm" : "text-muted-foreground"
              )}
              onClick={() => setActiveTab('plt-plh')}
            >
              <LayoutDashboard className="h-4 w-4" />
              PLT & PLH
            </Button>
          </nav>
        </div>

        <div className="mt-auto p-6 border-t">
          {user && (
            <div className="flex items-center gap-3 mb-4">
              <div className="h-8 w-8 rounded-full bg-primary/10 flex items-center justify-center">
                <User className="h-4 w-4 text-primary" />
              </div>
              <div className="flex-1 min-w-0">
                <p className="text-xs font-bold truncate">{user.displayName}</p>
                <p className="text-[10px] text-muted-foreground truncate">{user.email}</p>
              </div>
            </div>
          )}
          {user ? (
            <Button variant="outline" className="w-full justify-start gap-2 text-xs h-9" onClick={handleLogout}>
              <LogOut className="h-3.5 w-3.5" />
              Keluar
            </Button>
          ) : (
            <Button className="w-full justify-start gap-2 text-xs h-9" onClick={handleLogin}>
              <LogIn className="h-3.5 w-3.5" />
              Masuk
            </Button>
          )}
        </div>
      </aside>

      <div className="flex-1 flex flex-col min-w-0">
        {/* Header */}
        <header className="sticky top-0 z-40 w-full border-b bg-white/80 backdrop-blur-md h-16 flex items-center shrink-0">
          <div className="container mx-auto flex items-center justify-between px-4 md:px-8">
            <div className="md:hidden flex items-center gap-2">
              <span className="text-2xl font-black tracking-tighter text-[#1A4A9A]">BSKJI</span>
              <div className="h-6 w-[1px] bg-muted-foreground/20 mx-1" />
              <div className="flex flex-col">
                <p className="text-[7px] font-bold text-[#1A4A9A] leading-none uppercase">Badan Standardisasi dan</p>
                <p className="text-[7px] font-bold text-[#1A4A9A] leading-none uppercase">Kebijakan Jasa Industri</p>
              </div>
            </div>

            <div className="hidden md:block">
              <h2 className="text-sm font-semibold text-muted-foreground">
                {activeTab === 'tim-kerja' ? 'Monitoring Tim Kerja' : activeTab === 'skp' ? 'Monitoring SKP Pegawai' : 'Monitoring PLT & PLH'}
              </h2>
            </div>

            <div className="flex items-center gap-3">
              {user && activeTab === 'tim-kerja' && (
                <Dialog open={isAddTeamDialogOpen} onOpenChange={(open) => {
                  setIsAddTeamDialogOpen(open);
                  if (!open) {
                    setPendingSubTeams([]);
                    setJenisAddTeam('Tim Kerja');
                  }
                }}>
                  <DialogTrigger 
                    nativeButton={true}
                    render={<Button size="sm" className="gap-2 h-9 shadow-sm" />}
                  >
                    <Plus className="h-4 w-4" />
                    Tambah Tim Kerja
                  </DialogTrigger>
                  <DialogContent className={cn("max-h-[90vh] overflow-y-auto", jenisAddTeam === 'Bagian' ? "sm:max-w-[700px]" : "sm:max-w-[500px]")}>
                    <form onSubmit={handleAddTeam}>
                      <DialogHeader>
                        <DialogTitle>Tambah {jenisAddTeam === 'Bagian' ? 'Bagian' : 'Tim Kerja'} Baru</DialogTitle>
                        <DialogDescription>
                          Buat {jenisAddTeam === 'Bagian' ? 'bagian unit kerja' : 'tim kerja'} baru untuk monitoring penugasan.
                        </DialogDescription>
                      </DialogHeader>
                      <div className="grid gap-4 py-4 max-h-[65vh] overflow-y-auto pr-2 custom-scrollbar">
                        <div className="grid gap-2">
                          <Label htmlFor="jenis">Jenis</Label>
                          <Select name="jenis" defaultValue="Tim Kerja" onValueChange={(v: any) => {
                            setJenisAddTeam(v);
                            if (v === 'Tim Kerja') setPendingSubTeams([]);
                          }} required>
                            <SelectTrigger>
                              <SelectValue placeholder="Pilih Jenis" />
                            </SelectTrigger>
                            <SelectContent>
                              <SelectItem value="Tim Kerja">Tim Kerja</SelectItem>
                              <SelectItem value="Bagian">Bagian</SelectItem>
                            </SelectContent>
                          </Select>
                        </div>
                        {jenisAddTeam === 'Tim Kerja' && (
                          <div className="grid gap-2">
                            <Label htmlFor="parentId">Bagian (Opsional)</Label>
                            <Select name="parentId">
                              <SelectTrigger>
                                <SelectValue placeholder="Pilih Bagian" />
                              </SelectTrigger>
                              <SelectContent>
                                <SelectItem value="">Tanpa Bagian</SelectItem>
                                {workTeams.filter(t => t.jenis === 'Bagian').map(bagian => (
                                  <SelectItem key={bagian.id} value={bagian.id}>{bagian.namaTim}</SelectItem>
                                ))}
                              </SelectContent>
                            </Select>
                          </div>
                        )}
                        <div className="grid gap-2">
                          <Label htmlFor="namaTim">{jenisAddTeam === 'Bagian' ? 'Nama Bagian' : 'Nama Tim Kerja'}</Label>
                          <Input id="namaTim" name="namaTim" placeholder={jenisAddTeam === 'Bagian' ? 'Contoh: Bagian Tata Usaha' : 'Contoh: Tim Kerja Kepegawaian'} required />
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
                        <div className="grid grid-cols-2 gap-4">
                          <div className="grid gap-2">
                            <Label htmlFor="tahun">Tahun</Label>
                            <Input id="tahun" name="tahun" type="number" defaultValue={new Date().getFullYear()} required />
                          </div>
                        </div>

                        <div className="border-t pt-4 mt-2">
                          <Label className="text-primary font-bold">{jenisAddTeam === 'Bagian' ? 'Kepala Bagian' : 'Ketua Tim'}</Label>
                          <div className="grid gap-4 mt-2">
                            <div className="grid gap-2">
                              <Label htmlFor="ketuaNama">Nama {jenisAddTeam === 'Bagian' ? 'Kepala' : 'Ketua'}</Label>
                              <Input id="ketuaNama" name="ketuaNama" placeholder={`Nama lengkap ${jenisAddTeam === 'Bagian' ? 'kepala' : 'ketua'}`} required />
                            </div>
                            <div className="grid gap-2">
                              <Label htmlFor="ketuaJabatan">Jabatan {jenisAddTeam === 'Bagian' ? 'Kepala' : 'Ketua'}</Label>
                              <Input id="ketuaJabatan" name="ketuaJabatan" placeholder="Jabatan" required />
                            </div>
                            <div className="grid gap-2">
                              <Label htmlFor="ketuaStatus">Status {jenisAddTeam === 'Bagian' ? 'Kepala' : 'Ketua'}</Label>
                              <Select name="ketuaStatus" required>
                                <SelectTrigger>
                                  <SelectValue placeholder="Pilih Status" />
                                </SelectTrigger>
                                <SelectContent>
                                  {STATUS_PEGAWAI_LIST.map(status => (
                                    <SelectItem key={status} value={status}>{status}</SelectItem>
                                  ))}
                                </SelectContent>
                              </Select>
                            </div>
                          </div>
                        </div>

                        {jenisAddTeam === 'Bagian' && (
                          <div className="border-t pt-4 mt-2">
                            <div className="flex items-center justify-between mb-4">
                              <div className="flex flex-col">
                                <Label className="text-primary font-bold uppercase tracking-wider text-[10px]">Struktur Tim Kerja</Label>
                                <span className="text-[10px] text-muted-foreground uppercase">Tambah tim di bawah bagian ini</span>
                              </div>
                              <Button type="button" variant="outline" size="sm" onClick={addPendingSubTeam} className="gap-1 h-8 text-[10px] font-bold">
                                <Plus className="h-3.5 w-3.5" /> TAMBAH TIM KERJA
                              </Button>
                            </div>
                            
                            <div className="space-y-4">
                              {pendingSubTeams.length === 0 && (
                                <div className="text-center py-8 border border-dashed rounded-lg text-muted-foreground text-[10px] uppercase font-medium bg-muted/5">
                                  Belum ada tim kerja ditambahkan
                                </div>
                              )}
                              {pendingSubTeams.map((sub, sIndex) => (
                                <Card key={sIndex} className="p-4 border-dashed relative shadow-none border-primary/20">
                                  <Button 
                                    type="button" 
                                    variant="ghost" 
                                    size="icon" 
                                    className="absolute top-2 right-2 h-7 w-7 text-muted-foreground hover:text-destructive"
                                    onClick={() => removePendingSubTeam(sIndex)}
                                  >
                                    <X className="h-4 w-4" />
                                  </Button>
                                  
                                  <div className="grid gap-4">
                                    <div className="grid gap-2">
                                      <Label className="text-[10px] uppercase font-bold text-muted-foreground tracking-wider">Nama Tim Kerja</Label>
                                      <Input 
                                        placeholder="Contoh: Tim Kerja Monitoring" 
                                        className="h-9 text-sm focus-visible:ring-primary"
                                        value={sub.namaTim}
                                        onChange={(e) => updatePendingSubTeam(sIndex, 'namaTim', e.target.value)}
                                      />
                                    </div>
                                    
                                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                      <div className="grid gap-2">
                                        <Label className="text-[10px] uppercase font-bold text-muted-foreground tracking-wider">Kepala / Ketua Tim</Label>
                                        <Input 
                                          placeholder="Nama Lengkap" 
                                          className="h-9 text-sm"
                                          value={sub.ketua.nama}
                                          onChange={(e) => updatePendingSubTeam(sIndex, 'ketua.nama', e.target.value)}
                                        />
                                      </div>
                                      <div className="grid gap-2">
                                        <Label className="text-[10px] uppercase font-bold text-muted-foreground tracking-wider">Jabatan</Label>
                                        <Input 
                                          placeholder="Contoh: Ketua Tim Kerja" 
                                          className="h-9 text-sm"
                                          value={sub.ketua.jabatan}
                                          onChange={(e) => updatePendingSubTeam(sIndex, 'ketua.jabatan', e.target.value)}
                                        />
                                      </div>
                                    </div>

                                    <div className="grid gap-2">
                                      <Label className="text-[10px] uppercase font-bold text-muted-foreground tracking-wider">Status Kepegawaian Ketua</Label>
                                      <Select 
                                        value={sub.ketua.status}
                                        onValueChange={(v: any) => updatePendingSubTeam(sIndex, 'ketua.status', v)}
                                      >
                                        <SelectTrigger className="h-9 text-sm">
                                          <SelectValue />
                                        </SelectTrigger>
                                        <SelectContent>
                                          {STATUS_PEGAWAI_LIST.map(st => <SelectItem key={st} value={st}>{st}</SelectItem>)}
                                        </SelectContent>
                                      </Select>
                                    </div>
                                    
                                    <div className="border-t pt-4 mt-2">
                                      <div className="flex items-center justify-between mb-3">
                                        <span className="text-[10px] font-black text-muted-foreground uppercase tracking-widest">Anggota Tim ({sub.anggota.length})</span>
                                        <Button type="button" variant="ghost" size="sm" onClick={() => addMemberToPendingSubTeam(sIndex)} className="h-7 text-[10px] px-2 font-bold text-primary hover:bg-primary/5">
                                          + TAMBAH ANGGOTA
                                        </Button>
                                      </div>
                                      
                                      <div className="space-y-3">
                                        {sub.anggota.map((member, mIndex) => (
                                          <div key={mIndex} className="bg-muted/30 p-3 rounded border relative space-y-3">
                                            <Button 
                                              type="button" 
                                              variant="ghost" 
                                              size="icon" 
                                              className="absolute top-1 right-1 h-6 w-6 text-muted-foreground hover:text-destructive"
                                              onClick={() => removeMemberFromPendingSubTeam(sIndex, mIndex)}
                                            >
                                              <X className="h-3.5 w-3.5" />
                                            </Button>
                                            
                                            <div className="grid gap-2">
                                              <Label className="text-[9px] uppercase font-bold">Nama Anggota</Label>
                                              <Input 
                                                placeholder="Nama Lengkap" 
                                                className="h-8 text-xs"
                                                value={member.nama}
                                                onChange={(e) => updateMemberInPendingSubTeam(sIndex, mIndex, 'nama', e.target.value)}
                                              />
                                            </div>
                                            <div className="grid grid-cols-2 gap-2">
                                              <div className="grid gap-2">
                                                <Label className="text-[9px] uppercase font-bold">Jabatan dalam Tim</Label>
                                                <Input 
                                                  placeholder="Jabatan" 
                                                  className="h-8 text-xs"
                                                  value={member.jabatan}
                                                  onChange={(e) => updateMemberInPendingSubTeam(sIndex, mIndex, 'jabatan', e.target.value)}
                                                />
                                              </div>
                                              <div className="grid gap-2">
                                                <Label className="text-[9px] uppercase font-bold">Status</Label>
                                                <Select 
                                                  value={member.status} 
                                                  onValueChange={(v) => updateMemberInPendingSubTeam(sIndex, mIndex, 'status', v)}
                                                >
                                                  <SelectTrigger className="h-8 text-[10px]">
                                                    <SelectValue />
                                                  </SelectTrigger>
                                                  <SelectContent>
                                                    {STATUS_PEGAWAI_LIST.map(st => <SelectItem key={st} value={st}>{st}</SelectItem>)}
                                                  </SelectContent>
                                                </Select>
                                              </div>
                                            </div>
                                          </div>
                                        ))}
                                      </div>
                                    </div>
                                  </div>
                                </Card>
                              ))}
                            </div>
                          </div>
                        )}
                      </div>
                      <DialogFooter className="sticky bottom-0 bg-background pt-4 border-t z-10">
                        <DialogClose 
                          nativeButton={true}
                          render={<Button type="button" variant="outline">Batal</Button>} 
                        />
                        <Button type="submit" className="bg-primary hover:bg-primary/90 font-bold">
                          SIMPAN {jenisAddTeam.toUpperCase()}
                        </Button>
                      </DialogFooter>
                    </form>
                  </DialogContent>
                </Dialog>
              )}

              {user && activeTab === 'skp' && (
                <>
                  <Button 
                    variant="outline" 
                    size="sm" 
                    className="hidden sm:flex gap-2 h-9"
                    onClick={handleExport}
                  >
                    <Download className="h-4 w-4" />
                    Export Excel
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
                      render={<Button size="sm" className="gap-2 h-9 shadow-sm" />}
                    >
                      <Plus className="h-4 w-4" />
                      Tambah Data
                    </DialogTrigger>
                    <DialogContent className="sm:max-w-[550px] max-h-[90vh] overflow-y-auto">
                      {isAddDialogOpen && (
                        <form onSubmit={handleAddRecord}>
                          <DialogHeader>
                            <DialogTitle>Tambah Data SKP Baru</DialogTitle>
                            <DialogDescription>
                              Masukkan informasi pegawai untuk monitoring berkas SKP.
                            </DialogDescription>
                          </DialogHeader>
                          <div className="grid gap-4 py-4 max-h-[65vh] overflow-y-auto pr-2 custom-scrollbar">
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
                    )}
                  </DialogContent>
                </Dialog>
              </>
            )}

              {user && activeTab === 'plt-plh' && (
                <>
                  <Button 
                    variant="outline" 
                    size="sm" 
                    className="hidden sm:flex gap-2 h-9"
                    onClick={handleExportPltPlh}
                  >
                    <Download className="h-4 w-4" />
                    Export Excel
                  </Button>
                  <Dialog open={isAddPltPlhDialogOpen} onOpenChange={setIsAddPltPlhDialogOpen}>
                    <DialogTrigger 
                      nativeButton={true}
                      render={<Button size="sm" className="gap-2 h-9 shadow-sm" />}
                    >
                      <Plus className="h-4 w-4" />
                      Tambah PLT/PLH
                    </DialogTrigger>
                    <DialogContent className="sm:max-w-[550px] max-h-[90vh] overflow-y-auto">
                      {isAddPltPlhDialogOpen && (
                        <form onSubmit={handleAddPltPlhRecord}>
                        <DialogHeader>
                          <DialogTitle>Tambah Data PLT / PLH Baru</DialogTitle>
                          <DialogDescription>
                            Masukkan informasi penugasan PLT atau PLH pegawai.
                          </DialogDescription>
                        </DialogHeader>
                        <div className="grid gap-4 py-4 max-h-[65vh] overflow-y-auto pr-2 custom-scrollbar">
                          <div className="grid gap-2">
                            <Label>Jenis Penugasan</Label>
                            <Select 
                              name="type" 
                              defaultValue={pltPlhType} 
                              onValueChange={(value: 'PLT' | 'PLH') => setPltPlhType(value)}
                              required
                            >
                              <SelectTrigger>
                                <SelectValue placeholder="Pilih Jenis" />
                              </SelectTrigger>
                              <SelectContent>
                                <SelectItem value="PLT">PLT (Pelaksana Tugas)</SelectItem>
                                <SelectItem value="PLH">PLH (Pelaksana Harian)</SelectItem>
                              </SelectContent>
                            </Select>
                          </div>

                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div className="grid gap-2">
                              <Label htmlFor="namaPegawai">Nama Lengkap</Label>
                              <Input id="namaPegawai" name="namaPegawai" placeholder="Nama lengkap pegawai" required />
                            </div>
                            <div className="grid gap-2">
                              <Label htmlFor="nip">NIP</Label>
                              <Input id="nip" name="nip" placeholder="NIP pegawai" required />
                            </div>
                          </div>

                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div className="grid gap-2">
                              <Label htmlFor="jabatanAsli">Jabatan Asli</Label>
                              <Input id="jabatanAsli" name="jabatanAsli" placeholder="Jabatan definitif" required />
                            </div>
                            <div className="grid gap-2">
                              <Label htmlFor="jabatanTugas">Penugasan</Label>
                              <Input id="jabatanTugas" name="jabatanTugas" placeholder="Jabatan penugasan" required />
                            </div>
                          </div>

                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div className="grid gap-2">
                              <Label htmlFor="unitKerja">Unit Kerja Asal</Label>
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
                            <div className="grid gap-2">
                              <Label htmlFor="unitKerjaTugas">Unit Kerja Penugasan</Label>
                              <Select name="unitKerjaTugas" required>
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

                          <div className="grid gap-2">
                            <Label htmlFor="noSk">Nomor SK</Label>
                            <Input id="noSk" name="noSk" placeholder="Nomor Surat Keputusan" required />
                          </div>

                          <div className="grid gap-2">
                            <Label htmlFor="pegawaiDigantikan">Pegawai Asli / Sebelumnya yang Menjabat</Label>
                            <Input id="pegawaiDigantikan" name="pegawaiDigantikan" placeholder="Contoh: Budi Santoso (Pegawai yang digantikan)" />
                          </div>

                          <div className="grid gap-2">
                            <Label htmlFor="keterangan">Alasan / Keterangan Penugasan</Label>
                            <Input id="keterangan" name="keterangan" placeholder="Alasan penugasan PLT/PLH" />
                          </div>

                          <div className={cn("grid gap-4", pltPlhType === 'PLH' ? "grid-cols-1 sm:grid-cols-2" : "grid-cols-1")}>
                            <div className="grid gap-2">
                              <Label htmlFor="tglMulai">Tanggal Mulai</Label>
                              <Input id="tglMulai" name="tglMulai" type="date" required />
                            </div>
                            {pltPlhType === 'PLH' && (
                              <div className="grid gap-2">
                                <Label htmlFor="tglSelesai">Tanggal Selesai</Label>
                                <Input id="tglSelesai" name="tglSelesai" type="date" required />
                              </div>
                            )}
                          </div>
                        </div>
                        <DialogFooter className="gap-2 sm:gap-0 sticky bottom-0 bg-background pt-2 border-t mt-2">
                          <DialogClose 
                            nativeButton={true}
                            render={<Button type="button" variant="outline">Batal</Button>} 
                          />
                          <Button type="submit" className="bg-primary hover:bg-primary/90">
                            Simpan Data Penugasan
                          </Button>
                        </DialogFooter>
                      </form>
                    )}
                  </DialogContent>
                </Dialog>
              </>
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
                  {isEditDialogOpen && editingRecord && (
                    <form onSubmit={handleUpdateRecord}>
                      <DialogHeader>
                        <DialogTitle>Edit Data SKP</DialogTitle>
                        <DialogDescription>
                          Perbarui informasi pegawai jika terjadi mutasi atau perubahan data lainnya.
                        </DialogDescription>
                      </DialogHeader>
                      <div className="grid gap-4 py-4 max-h-[65vh] overflow-y-auto pr-2 custom-scrollbar">
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

        <main className="flex-1 overflow-y-auto p-4 md:p-8 space-y-8">
        {activeTab === 'tim-kerja' ? (
          <div className="space-y-6">
            {Object.keys(groupedTeams).length > 0 ? (
              (Object.entries(groupedTeams) as [string, { topLevel: WorkTeam[], subTeams: Record<string, WorkTeam[]> }][]).map(([unitKerja, data]) => (
                <div key={unitKerja} className="space-y-4">
                  <div className="flex items-center gap-2">
                    <Building2 className="h-5 w-5 text-primary" />
                    <h3 className="text-lg font-bold text-[#1A4A9A]">{unitKerja}</h3>
                    <Badge variant="outline" className="ml-2 bg-primary/5 text-primary border-primary/20">
                      {data.topLevel.length + Object.values(data.subTeams).flat().length} Tim
                    </Badge>
                  </div>
                  
                  <div className="space-y-6">
                    {data.topLevel.map((team) => (
                      <div key={team.id} className="space-y-4">
                        <div className="grid gap-6 md:grid-cols-2 lg:grid-cols-3">
                          {renderTeamCard(team)}
                        </div>
                        {data.subTeams[team.id] && data.subTeams[team.id].length > 0 && (
                          <div className="ml-8 pl-4 border-l-2 border-primary/10 space-y-4">
                            <div className="flex items-center gap-2 text-[10px] font-bold text-primary uppercase tracking-widest">
                              <ArrowRight className="h-3 w-3" />
                              Tim Kerja di bawah {team.namaTim}
                            </div>
                            <div className="grid gap-6 md:grid-cols-2 lg:grid-cols-3">
                              {data.subTeams[team.id].map(subTeam => renderTeamCard(subTeam, true))}
                            </div>
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              ))
            ) : (
              <div className="flex flex-col items-center justify-center py-20 bg-white rounded-2xl border border-dashed shadow-sm">
                <div className="h-16 w-16 rounded-full bg-primary/5 flex items-center justify-center mb-4">
                  <Users className="h-8 w-8 text-primary/40" />
                </div>
                <h3 className="text-lg font-bold text-muted-foreground">Belum Ada Tim Kerja</h3>
                <p className="text-sm text-muted-foreground/60 max-w-xs text-center mt-1">
                  Mulai dengan menambahkan tim kerja baru untuk memantau penugasan di setiap unit kerja.
                </p>
                <Button 
                  variant="outline" 
                  className="mt-6 gap-2"
                  onClick={() => setIsAddTeamDialogOpen(true)}
                >
                  <Plus className="h-4 w-4" />
                  Tambah Tim Pertama
                </Button>
              </div>
            )}
          </div>
        ) : activeTab === 'skp' ? (
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
                      {isAuthLoading ? (
                        <TableRow>
                          <TableCell colSpan={6} className="h-32 text-center text-muted-foreground">
                            Memuat data...
                          </TableCell>
                        </TableRow>
                      ) : !user ? (
                        <TableRow>
                          <TableCell colSpan={6} className="h-32 text-center">
                            <div className="flex flex-col items-center gap-2">
                              <p className="text-muted-foreground">Silakan masuk untuk melihat data.</p>
                              <Button size="sm" onClick={handleLogin}>Masuk dengan Google</Button>
                            </div>
                          </TableCell>
                        </TableRow>
                      ) : (
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
                                  <div className="flex items-center justify-end gap-1">
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
                                      <Button 
                                        size="icon" 
                                        variant="ghost" 
                                        className="h-8 w-8 opacity-0 group-hover:opacity-100 transition-opacity"
                                        onClick={() => {
                                          setEditingRecord(record);
                                          setIsEditDialogOpen(true);
                                        }}
                                      >
                                        <UserCog className="h-4 w-4" />
                                      </Button>
                                    )}
                                    <Button 
                                      size="icon" 
                                      variant="ghost" 
                                      className="h-8 w-8 text-destructive hover:text-destructive hover:bg-destructive/10 opacity-0 group-hover:opacity-100 transition-opacity"
                                      onClick={() => handleDeleteRecord(record.id)}
                                    >
                                      <X className="h-4 w-4" />
                                    </Button>
                                  </div>
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
                      )}
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
                <Card 
                  className={cn(
                    "overflow-hidden border-none shadow-sm hover:shadow-md transition-all cursor-pointer",
                    statusFilterPlt === 'all' ? "ring-2 ring-primary ring-offset-2" : ""
                  )}
                  onClick={() => setStatusFilterPlt('all')}
                >
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
                <Card 
                  className={cn(
                    "overflow-hidden border-none shadow-sm hover:shadow-md transition-all cursor-pointer",
                    statusFilterPlt === 'Aktif' ? "ring-2 ring-emerald-500 ring-offset-2" : ""
                  )}
                  onClick={() => setStatusFilterPlt(statusFilterPlt === 'Aktif' ? 'all' : 'Aktif')}
                >
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
                <Card 
                  className={cn(
                    "overflow-hidden border-none shadow-sm hover:shadow-md transition-all cursor-pointer",
                    statusFilterPlh === 'all' ? "ring-2 ring-blue-500 ring-offset-2" : ""
                  )}
                  onClick={() => setStatusFilterPlh('all')}
                >
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
                <Card 
                  className={cn(
                    "overflow-hidden border-none shadow-sm hover:shadow-md transition-all cursor-pointer",
                    statusFilterPlh === 'Aktif' ? "ring-2 ring-blue-400 ring-offset-2" : ""
                  )}
                  onClick={() => setStatusFilterPlh(statusFilterPlh === 'Aktif' ? 'all' : 'Aktif')}
                >
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

            {/* PLT/PLH Chart */}
            <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.5 }}>
              <Card className="border-none shadow-sm">
                <CardHeader>
                  <CardTitle className="text-lg">Grafik Penugasan PLT & PLH per Tahun</CardTitle>
                  <CardDescription>Perbandingan jumlah penugasan Pelaksana Tugas dan Pelaksana Harian.</CardDescription>
                </CardHeader>
                <CardContent>
                  <div className="h-[300px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart
                        data={chartDataPltPlh}
                        margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
                      >
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
                        <XAxis 
                          dataKey="year" 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ fill: '#666', fontSize: 12 }}
                          dy={10}
                        />
                        <YAxis 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ fill: '#666', fontSize: 12 }}
                        />
                        <Tooltip 
                          cursor={{ fill: '#f8f9fa' }}
                          contentStyle={{ 
                            borderRadius: '8px', 
                            border: 'none', 
                            boxShadow: '0 4px 12px rgba(0,0,0,0.1)' 
                          }}
                        />
                        <Legend verticalAlign="top" align="right" iconType="circle" wrapperStyle={{ paddingBottom: '20px' }} />
                        <Bar 
                          dataKey="PLT" 
                          fill="#1A4A9A" 
                          radius={[4, 4, 0, 0]} 
                          barSize={30}
                        />
                        <Bar 
                          dataKey="PLH" 
                          fill="#3b82f6" 
                          radius={[4, 4, 0, 0]} 
                          barSize={30}
                        />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </CardContent>
              </Card>
            </motion.div>

            {/* PLT Table */}
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-white border-b pb-6">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div>
                    <CardTitle>Daftar Pelaksana Tugas (PLT)</CardTitle>
                    <CardDescription>Monitoring penugasan Pelaksana Tugas di lingkungan BSKJI.</CardDescription>
                  </div>
                  <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
                    <div className="relative w-full md:w-64">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
                      <Input 
                        placeholder="Cari PLT..." 
                        className="pl-9 bg-muted/50 border-none focus-visible:ring-1"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                      />
                    </div>
                    <Select value={yearFilterPltPlh} onValueChange={setYearFilterPltPlh}>
                      <SelectTrigger className="w-full md:w-[130px] bg-muted/50 border-none h-10">
                        <div className="flex items-center gap-2">
                          <Filter className="h-3.5 w-3.5 text-muted-foreground" />
                          <SelectValue placeholder="Tahun" />
                        </div>
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="all">Semua Tahun</SelectItem>
                        {Array.from({ length: 5 }, (_, i) => new Date().getFullYear() - i).map(year => (
                          <SelectItem key={year} value={year.toString()}>{year}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    <Select value={statusFilterPlt} onValueChange={setStatusFilterPlt}>
                      <SelectTrigger className="w-full md:w-[130px] bg-muted/50 border-none h-10">
                        <div className="flex items-center gap-2">
                          <CheckCircle2 className="h-3.5 w-3.5 text-muted-foreground" />
                          <SelectValue placeholder="Status" />
                        </div>
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="all">Semua Status</SelectItem>
                        <SelectItem value="Aktif">Aktif</SelectItem>
                        <SelectItem value="Selesai">Selesai</SelectItem>
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
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Pegawai</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Penugasan</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Unit Kerja (Asal/Tugas)</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Periode</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Status</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider text-right">Aksi</TableHead>
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
                                {record.pegawaiDigantikan && (
                                  <span className="text-[10px] text-primary italic font-medium">Gantikan: {record.pegawaiDigantikan}</span>
                                )}
                              </div>
                            </TableCell>
                            <TableCell>
                              <div className="flex flex-col">
                                <span className="font-medium text-sm text-primary">{record.jabatanPlt}</span>
                                <span className="text-[10px] text-muted-foreground font-medium">SK: {record.noSk}</span>
                                {record.keterangan && (
                                  <span className="text-[10px] text-muted-foreground mt-1 bg-muted/50 px-1.5 py-0.5 rounded border border-muted-foreground/10">
                                    Ket: {record.keterangan}
                                  </span>
                                )}
                              </div>
                            </TableCell>
                            <TableCell>
                              <div className="flex flex-col gap-1">
                                <div className="flex items-center gap-2 text-sm">
                                  <Building2 className="h-3.5 w-3.5 text-muted-foreground" />
                                  <span className="font-medium">Asal:</span> {record.unitKerja}
                                </div>
                                <div className="flex items-center gap-2 text-sm">
                                  <ArrowRight className="h-3.5 w-3.5 text-primary" />
                                  <span className="font-medium">Tugas:</span> {record.unitKerjaTugas}
                                </div>
                              </div>
                            </TableCell>
                            <TableCell>
                              <div className="flex flex-col text-xs">
                                <span>{format(new Date(record.tglMulai), 'dd MMM yyyy')}</span>
                                <span className="text-muted-foreground">
                                  {record.tglSelesai 
                                    ? `s.d. ${format(new Date(record.tglSelesai), 'dd MMM yyyy')}` 
                                    : 's.d. Selesai'}
                                </span>
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
                            <TableCell className="text-right">
                              <div className="flex justify-end gap-2">
                                <Button 
                                  size="icon" 
                                  variant="ghost" 
                                  className="h-8 w-8 opacity-0 group-hover:opacity-100 transition-opacity"
                                  onClick={() => {
                                    setEditingPltRecord(record);
                                    setIsEditPltPlhDialogOpen(true);
                                  }}
                                >
                                  <UserCog className="h-4 w-4" />
                                </Button>
                                <Button 
                                  size="icon" 
                                  variant="ghost" 
                                  className="h-8 w-8 text-destructive hover:text-destructive hover:bg-destructive/10 opacity-0 group-hover:opacity-100 transition-opacity"
                                  onClick={() => handleDeletePltPlhRecord(record.id, 'PLT')}
                                >
                                  <X className="h-4 w-4" />
                                </Button>
                              </div>
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
                  <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
                    <div className="relative w-full md:w-64">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
                      <Input 
                        placeholder="Cari PLH..." 
                        className="pl-9 bg-muted/50 border-none focus-visible:ring-1"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                      />
                    </div>
                    <Select value={yearFilterPltPlh} onValueChange={setYearFilterPltPlh}>
                      <SelectTrigger className="w-full md:w-[130px] bg-muted/50 border-none h-10">
                        <div className="flex items-center gap-2">
                          <Filter className="h-3.5 w-3.5 text-muted-foreground" />
                          <SelectValue placeholder="Tahun" />
                        </div>
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="all">Semua Tahun</SelectItem>
                        {Array.from({ length: 5 }, (_, i) => new Date().getFullYear() - i).map(year => (
                          <SelectItem key={year} value={year.toString()}>{year}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    <Select value={statusFilterPlh} onValueChange={setStatusFilterPlh}>
                      <SelectTrigger className="w-full md:w-[130px] bg-muted/50 border-none h-10">
                        <div className="flex items-center gap-2">
                          <CheckCircle2 className="h-3.5 w-3.5 text-muted-foreground" />
                          <SelectValue placeholder="Status" />
                        </div>
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="all">Semua Status</SelectItem>
                        <SelectItem value="Aktif">Aktif</SelectItem>
                        <SelectItem value="Selesai">Selesai</SelectItem>
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
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Pegawai</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Penugasan</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Unit Kerja (Asal/Tugas)</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Periode</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider">Status</TableHead>
                        <TableHead className="font-bold text-xs uppercase tracking-wider text-right">Aksi</TableHead>
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
                                {record.pegawaiDigantikan && (
                                  <span className="text-[10px] text-blue-600 italic font-medium">Gantikan: {record.pegawaiDigantikan}</span>
                                )}
                              </div>
                            </TableCell>
                            <TableCell>
                              <div className="flex flex-col">
                                <span className="font-medium text-sm text-blue-600">{record.jabatanPlh}</span>
                                <span className="text-[10px] text-muted-foreground font-medium">SK: {record.noSk}</span>
                                {record.keterangan && (
                                  <span className="text-[10px] text-muted-foreground mt-1 bg-muted/50 px-1.5 py-0.5 rounded border border-muted-foreground/10">
                                    Ket: {record.keterangan}
                                  </span>
                                )}
                              </div>
                            </TableCell>
                            <TableCell>
                              <div className="flex flex-col gap-1">
                                <div className="flex items-center gap-2 text-sm">
                                  <Building2 className="h-3.5 w-3.5 text-muted-foreground" />
                                  <span className="font-medium">Asal:</span> {record.unitKerja}
                                </div>
                                <div className="flex items-center gap-2 text-sm">
                                  <ArrowRight className="h-3.5 w-3.5 text-primary" />
                                  <span className="font-medium">Tugas:</span> {record.unitKerjaTugas}
                                </div>
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
                            <TableCell className="text-right">
                              <div className="flex justify-end gap-2">
                                <Button 
                                  size="icon" 
                                  variant="ghost" 
                                  className="h-8 w-8 opacity-0 group-hover:opacity-100 transition-opacity"
                                  onClick={() => {
                                    setEditingPlhRecord(record);
                                    setIsEditPltPlhDialogOpen(true);
                                  }}
                                >
                                  <UserCog className="h-4 w-4" />
                                </Button>
                                <Button 
                                  size="icon" 
                                  variant="ghost" 
                                  className="h-8 w-8 text-destructive hover:text-destructive hover:bg-destructive/10 opacity-0 group-hover:opacity-100 transition-opacity"
                                  onClick={() => handleDeletePltPlhRecord(record.id, 'PLH')}
                                >
                                  <X className="h-4 w-4" />
                                </Button>
                              </div>
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

      {/* Edit PLT/PLH Dialog */}
      <Dialog open={isEditPltPlhDialogOpen} onOpenChange={(open) => {
        setIsEditPltPlhDialogOpen(open);
        if (!open) {
          setEditingPltRecord(null);
          setEditingPlhRecord(null);
        }
      }}>
        <DialogContent className="sm:max-w-[500px] max-h-[90vh] overflow-y-auto">
          {isEditPltPlhDialogOpen && (
            <form onSubmit={handleUpdatePltPlhRecord}>
              <DialogHeader>
                <DialogTitle>Edit Data {editingPltRecord ? 'PLT' : 'PLH'}</DialogTitle>
                <DialogDescription>
                  Perbarui informasi penugasan pegawai.
                </DialogDescription>
              </DialogHeader>
              <div className="grid gap-6 py-6 max-h-[65vh] overflow-y-auto pr-2 custom-scrollbar">
                <input type="hidden" name="type" value={editingPltRecord ? 'PLT' : 'PLH'} />
                
                <div className="grid gap-2">
                  <Label htmlFor="edit-namaPegawai">Nama Lengkap</Label>
                  <Input 
                    id="edit-namaPegawai" 
                    name="namaPegawai" 
                    defaultValue={editingPltRecord?.namaPegawai || editingPlhRecord?.namaPegawai} 
                    required 
                  />
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="grid gap-2">
                    <Label htmlFor="edit-nip">NIP</Label>
                    <Input 
                      id="edit-nip" 
                      name="nip" 
                      defaultValue={editingPltRecord?.nip || editingPlhRecord?.nip} 
                      required 
                    />
                  </div>
                  <div className="grid gap-2">
                    <Label htmlFor="edit-status">Status</Label>
                    <Select name="status" defaultValue={editingPltRecord?.status || editingPlhRecord?.status || 'Aktif'}>
                      <SelectTrigger>
                        <SelectValue placeholder="Pilih Status" />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="Aktif">Aktif</SelectItem>
                        <SelectItem value="Selesai">Selesai</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="grid gap-2">
                    <Label htmlFor="edit-jabatanAsli">Jabatan Asli</Label>
                    <Input 
                      id="edit-jabatanAsli" 
                      name="jabatanAsli" 
                      defaultValue={editingPltRecord?.jabatanAsli || editingPlhRecord?.jabatanAsli} 
                      required 
                    />
                  </div>
                  <div className="grid gap-2">
                    <Label htmlFor="edit-jabatanTugas">Penugasan</Label>
                    <Input 
                      id="edit-jabatanTugas" 
                      name="jabatanTugas" 
                      defaultValue={editingPltRecord?.jabatanPlt || editingPlhRecord?.jabatanPlh} 
                      required 
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="grid gap-2">
                    <Label htmlFor="edit-unitKerja">Unit Kerja Asal</Label>
                    <Select name="unitKerja" defaultValue={editingPltRecord?.unitKerja || editingPlhRecord?.unitKerja}>
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
                    <Label htmlFor="edit-unitKerjaTugas">Unit Kerja Penugasan</Label>
                    <Select name="unitKerjaTugas" defaultValue={editingPltRecord?.unitKerjaTugas || editingPlhRecord?.unitKerjaTugas}>
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

                <div className="grid gap-2">
                  <Label htmlFor="edit-noSk">Nomor SK</Label>
                  <Input 
                    id="edit-noSk" 
                    name="noSk" 
                    defaultValue={editingPltRecord?.noSk || editingPlhRecord?.noSk} 
                    required 
                  />
                </div>

                <div className="grid gap-2">
                  <Label htmlFor="edit-pegawaiDigantikan">Pegawai Asli / Sebelumnya yang Menjabat</Label>
                  <Input 
                    id="edit-pegawaiDigantikan" 
                    name="edit-pegawaiDigantikan" 
                    defaultValue={editingPltRecord?.pegawaiDigantikan || editingPlhRecord?.pegawaiDigantikan} 
                    placeholder="Nama pegawai asli"
                  />
                </div>

                <div className="grid gap-2">
                  <Label htmlFor="edit-keterangan">Alasan / Keterangan Penugasan</Label>
                  <Input 
                    id="edit-keterangan" 
                    name="keterangan" 
                    defaultValue={editingPltRecord?.keterangan || editingPlhRecord?.keterangan} 
                    placeholder="Alasan penugasan PLT/PLH"
                  />
                </div>

                <div className={cn("grid gap-4", editingPlhRecord ? "grid-cols-1 sm:grid-cols-2" : "grid-cols-1")}>
                  <div className="grid gap-2">
                    <Label htmlFor="edit-tglMulai">Tanggal Mulai</Label>
                    <Input 
                      id="edit-tglMulai" 
                      name="tglMulai" 
                      type="date" 
                      defaultValue={editingPltRecord?.tglMulai || editingPlhRecord?.tglMulai} 
                      required 
                    />
                  </div>
                  {editingPlhRecord && (
                    <div className="grid gap-2">
                      <Label htmlFor="edit-tglSelesai">Tanggal Selesai</Label>
                      <Input 
                        id="edit-tglSelesai" 
                        name="tglSelesai" 
                        type="date" 
                        defaultValue={editingPlhRecord?.tglSelesai} 
                        required 
                      />
                    </div>
                  )}
                </div>
              </div>
              <DialogFooter>
                <DialogClose 
                  nativeButton={true}
                  render={<Button type="button" variant="outline">Batal</Button>} 
                />
                <Button type="submit">Simpan Perubahan</Button>
              </DialogFooter>
            </form>
          )}
        </DialogContent>
      </Dialog>

      {/* Take SKP Dialog */}
      <Dialog open={isTakeDialogOpen} onOpenChange={setIsTakeDialogOpen}>
        <DialogContent className="sm:max-w-[425px] max-h-[90vh] overflow-y-auto">
          {isTakeDialogOpen && (
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
                  render={<Button type="button" variant="ghost">Batal</Button>} 
                />
                <Button type="submit" className="bg-emerald-600 hover:bg-emerald-700">Konfirmasi Ambil</Button>
              </DialogFooter>
            </form>
          )}
        </DialogContent>
      </Dialog>

      {/* Detail Dialog */}
      <Dialog open={isDetailDialogOpen} onOpenChange={setIsDetailDialogOpen}>
        <DialogContent className="sm:max-w-[500px] max-h-[90vh] overflow-y-auto">
          {isDetailDialogOpen && detailRecord && (
            <>
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
          </>
          )}
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

      {/* Edit Team Dialog */}
      <Dialog open={isEditTeamDialogOpen} onOpenChange={(open) => {
        setIsEditTeamDialogOpen(open);
        if (!open) setEditingTeam(null);
      }}>
        <DialogContent className="sm:max-w-[500px]">
          {editingTeam && (
            <form onSubmit={handleUpdateTeam}>
              <DialogHeader>
                <DialogTitle>Edit Tim Kerja</DialogTitle>
                <DialogDescription>
                  Perbarui informasi tim kerja dan ketua tim.
                </DialogDescription>
              </DialogHeader>
              <div className="grid gap-4 py-4 max-h-[65vh] overflow-y-auto pr-2 custom-scrollbar">
                <div className="grid gap-2">
                  <Label htmlFor="edit-jenis">Jenis</Label>
                  <Select name="jenis" defaultValue={editingTeam.jenis || 'Tim Kerja'} onValueChange={(v: any) => setJenisEditTeam(v)} required>
                    <SelectTrigger>
                      <SelectValue placeholder="Pilih Jenis" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="Tim Kerja">Tim Kerja</SelectItem>
                      <SelectItem value="Bagian">Bagian</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
                {jenisEditTeam === 'Tim Kerja' && (
                  <div className="grid gap-2">
                    <Label htmlFor="edit-parentId">Bagian (Opsional)</Label>
                    <Select name="parentId" defaultValue={editingTeam.parentId || ""}>
                      <SelectTrigger>
                        <SelectValue placeholder="Pilih Bagian" />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="">Tanpa Bagian</SelectItem>
                        {workTeams.filter(t => t.jenis === 'Bagian' && t.id !== editingTeam.id).map(bagian => (
                          <SelectItem key={bagian.id} value={bagian.id}>{bagian.namaTim}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                )}
                <div className="grid gap-2">
                  <Label htmlFor="edit-namaTim">{jenisEditTeam === 'Bagian' ? 'Nama Bagian' : 'Nama Tim Kerja'}</Label>
                  <Input id="edit-namaTim" name="namaTim" defaultValue={editingTeam.namaTim} required />
                </div>
                <div className="grid gap-2">
                  <Label htmlFor="edit-unitKerja">Unit Kerja</Label>
                  <Select name="unitKerja" defaultValue={editingTeam.unitKerja} required>
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
                <div className="grid grid-cols-2 gap-4">
                  <div className="grid gap-2">
                    <Label htmlFor="edit-tahun">Tahun</Label>
                    <Input id="edit-tahun" name="tahun" type="number" defaultValue={editingTeam.tahun} required />
                  </div>
                  <div className="grid gap-2">
                    <Label htmlFor="edit-status">Status</Label>
                    <Select name="status" defaultValue={editingTeam.status} required>
                      <SelectTrigger>
                        <SelectValue placeholder="Pilih Status" />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="Aktif">Aktif</SelectItem>
                        <SelectItem value="Non-Aktif">Non-Aktif</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                </div>
                <div className="border-t pt-4 mt-2">
                  <Label className="text-primary font-bold">Ketua Tim</Label>
                  <div className="grid gap-4 mt-2">
                    <div className="grid gap-2">
                      <Label htmlFor="edit-ketuaNama">Nama Ketua</Label>
                      <Input id="edit-ketuaNama" name="ketuaNama" defaultValue={editingTeam.ketua.nama} required />
                    </div>
                    <div className="grid gap-2">
                      <Label htmlFor="edit-ketuaJabatan">Jabatan Ketua</Label>
                      <Input id="edit-ketuaJabatan" name="ketuaJabatan" defaultValue={editingTeam.ketua.jabatan} required />
                    </div>
                    <div className="grid gap-2">
                      <Label htmlFor="edit-ketuaStatus">Status Ketua</Label>
                      <Select name="ketuaStatus" defaultValue={editingTeam.ketua.status} required>
                        <SelectTrigger>
                          <SelectValue placeholder="Pilih Status" />
                        </SelectTrigger>
                        <SelectContent>
                          {STATUS_PEGAWAI_LIST.map(status => (
                            <SelectItem key={status} value={status}>{status}</SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </div>
                  </div>
                </div>
              </div>
              <DialogFooter>
                <DialogClose 
                  nativeButton={true}
                  render={<Button type="button" variant="outline">Batal</Button>} 
                />
                <Button type="submit">Simpan Perubahan</Button>
              </DialogFooter>
            </form>
          )}
        </DialogContent>
      </Dialog>

      {/* Add Member Dialog */}
      <Dialog open={isAddMemberDialogOpen} onOpenChange={(open) => {
        setIsAddMemberDialogOpen(open);
        if (!open) setSelectedTeamForMember(null);
      }}>
        <DialogContent className="sm:max-w-[400px]">
          <form onSubmit={(e) => {
            e.preventDefault();
            if (!selectedTeamForMember) return;
            const formData = new FormData(e.currentTarget);
            const member: TeamMember = {
              nama: formData.get('nama') as string,
              jabatan: formData.get('jabatan') as string,
              status: formData.get('status') as string,
            };
            handleAddMember(selectedTeamForMember, member);
            setIsAddMemberDialogOpen(false);
          }}>
            <DialogHeader>
              <DialogTitle>Tambah Anggota Tim</DialogTitle>
              <DialogDescription>
                Masukkan informasi anggota tim kerja baru.
              </DialogDescription>
            </DialogHeader>
            <div className="grid gap-4 py-4 max-h-[65vh] overflow-y-auto pr-2 custom-scrollbar">
              <div className="grid gap-2">
                <Label htmlFor="member-nama">Nama Anggota</Label>
                <Input id="member-nama" name="nama" placeholder="Nama lengkap anggota" required />
              </div>
              <div className="grid gap-2">
                <Label htmlFor="member-jabatan">Jabatan Anggota</Label>
                <Input id="member-jabatan" name="jabatan" placeholder="Jabatan anggota" required />
              </div>
              <div className="grid gap-2">
                <Label htmlFor="member-status">Status</Label>
                <Select name="status" defaultValue="PNS" required>
                  <SelectTrigger>
                    <SelectValue placeholder="Pilih Status" />
                  </SelectTrigger>
                  <SelectContent>
                    {STATUS_PEGAWAI_LIST.map(status => (
                      <SelectItem key={status} value={status}>{status}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>
            <DialogFooter>
              <DialogClose 
                nativeButton={true}
                render={<Button type="button" variant="outline">Batal</Button>} 
              />
              <Button type="submit">Tambah Anggota</Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>

      {/* Edit Member Dialog */}
      <Dialog open={isEditMemberDialogOpen} onOpenChange={(open) => {
        setIsEditMemberDialogOpen(open);
        if (!open) setEditingMemberInfo(null);
      }}>
        <DialogContent className="sm:max-w-[400px]">
          {editingMemberInfo && (
            <form onSubmit={(e) => {
              e.preventDefault();
              const formData = new FormData(e.currentTarget);
              const updatedMember: TeamMember = {
                nama: formData.get('nama') as string,
                jabatan: formData.get('jabatan') as string,
                status: formData.get('status') as string,
              };
              handleUpdateMember(editingMemberInfo.teamId, editingMemberInfo.index, updatedMember);
              setIsEditMemberDialogOpen(false);
            }}>
              <DialogHeader>
                <DialogTitle>Edit Anggota Tim</DialogTitle>
                <DialogDescription>
                  Perbarui informasi anggota tim kerja.
                </DialogDescription>
              </DialogHeader>
              <div className="grid gap-4 py-4 max-h-[65vh] overflow-y-auto pr-2 custom-scrollbar">
                <div className="grid gap-2">
                  <Label htmlFor="edit-member-nama">Nama Anggota</Label>
                  <Input id="edit-member-nama" name="nama" defaultValue={editingMemberInfo.member.nama} required />
                </div>
                <div className="grid gap-2">
                  <Label htmlFor="edit-member-jabatan">Jabatan Anggota</Label>
                  <Input id="edit-member-jabatan" name="jabatan" defaultValue={editingMemberInfo.member.jabatan} required />
                </div>
                <div className="grid gap-2">
                  <Label htmlFor="edit-member-status">Status</Label>
                  <Select name="status" defaultValue={editingMemberInfo.member.status} required>
                    <SelectTrigger>
                      <SelectValue placeholder="Pilih Status" />
                    </SelectTrigger>
                    <SelectContent>
                      {STATUS_PEGAWAI_LIST.map(status => (
                        <SelectItem key={status} value={status}>{status}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </div>
              <DialogFooter>
                <DialogClose 
                  nativeButton={true}
                  render={<Button type="button" variant="outline">Batal</Button>} 
                />
                <Button type="submit">Simpan Perubahan</Button>
              </DialogFooter>
            </form>
          )}
        </DialogContent>
      </Dialog>

      {/* Delete Confirmation Dialog */}
      <Dialog open={isDeleteConfirmOpen} onOpenChange={setIsDeleteConfirmOpen}>
        <DialogContent className="sm:max-w-[400px]">
          <DialogHeader>
            <div className="flex items-center gap-3 text-destructive mb-2">
              <div className="h-10 w-10 rounded-full bg-destructive/10 flex items-center justify-center">
                <AlertTriangle className="h-6 w-6" />
              </div>
              <DialogTitle>Konfirmasi Hapus</DialogTitle>
            </div>
            <DialogDescription>
              Apakah Anda yakin ingin menghapus data {deletingInfo?.type === 'SKP' ? 'SKP' : (deletingInfo?.type === 'TIM' ? 'Tim Kerja' : deletingInfo?.type)} ini? Tindakan ini tidak dapat dibatalkan.
            </DialogDescription>
          </DialogHeader>
          <DialogFooter className="mt-4 gap-2 sm:gap-0">
            <DialogClose 
              nativeButton={true}
              render={<Button variant="outline">Batal</Button>} 
            />
            <Button 
              variant="destructive" 
              onClick={confirmDelete}
            >
              Hapus Sekarang
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <footer className="container mx-auto p-8 text-center text-xs text-muted-foreground mb-16 md:mb-0">
        <p>&copy; {new Date().getFullYear()} Monitoring SKP Pegawai. Sistem Monitoring Berkas Kepegawaian.</p>
      </footer>

      {/* Mobile Navigation */}
      <div className="md:hidden fixed bottom-0 left-0 right-0 bg-white border-t z-50 flex items-center justify-around h-16 px-4 shadow-[0_-2px_10px_rgba(0,0,0,0.05)]">
        <Button 
          variant="ghost" 
          className={cn("flex flex-col items-center gap-1 h-auto py-1 px-0 flex-1", activeTab === 'tim-kerja' ? "text-primary" : "text-muted-foreground")}
          onClick={() => setActiveTab('tim-kerja')}
        >
          <Users className="h-5 w-5" />
          <span className="text-[10px] font-bold">Tim Kerja</span>
        </Button>
        <Button 
          variant="ghost" 
          className={cn("flex flex-col items-center gap-1 h-auto py-1 px-0 flex-1", activeTab === 'skp' ? "text-primary" : "text-muted-foreground")}
          onClick={() => setActiveTab('skp')}
        >
          <FileText className="h-5 w-5" />
          <span className="text-[10px] font-bold">SKP</span>
        </Button>
        <Button 
          variant="ghost" 
          className={cn("flex flex-col items-center gap-1 h-auto py-1 px-0 flex-1", activeTab === 'plt-plh' ? "text-primary" : "text-muted-foreground")}
          onClick={() => setActiveTab('plt-plh')}
        >
          <LayoutDashboard className="h-5 w-5" />
          <span className="text-[10px] font-bold">PLT/PLH</span>
        </Button>
      </div>
    </div>
  </div>
);
}
