/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import * as React from 'react';
import { useState, useMemo, useRef } from 'react';
import { 
  Search, 
  Plus, 
  FileCheck, 
  FileClock, 
  Users, 
  LayoutDashboard,
  Filter,
  Download,
  Upload,
  MoreVertical,
  CheckCircle2,
  XCircle,
  Calendar,
  Briefcase,
  Building2,
  ArrowRight,
  ArrowLeft,
  User,
  FileText,
  X,
  UserCheck,
  UserCog,
  LogOut,
  LogIn,
  AlertTriangle,
  Pencil,
  Archive,
  ExternalLink,
  Network,
  Trash2,
  TrendingUp,
  Clock
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

import { SKPRecord, SKPMonitoringRecord, SKPStatus, SKPPeriode, PLTRecord, PLHRecord, PLTStatus, WorkTeam, TeamMember, UnitInfo, OrgNode } from './types';

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

function OrgNodeCard({ node, allNodes, onAddSub, onEdit, onDelete }: { 
  node: OrgNode, 
  allNodes: OrgNode[], 
  onAddSub: (id: string) => void,
  onEdit: (node: OrgNode) => void,
  onDelete: (id: string) => void
}) {
  const children = allNodes.filter(n => n.parentId === node.id);
  const isRoot = !node.parentId;

  return (
    <div className="flex flex-col items-center relative">
      <motion.div
        initial={{ scale: 0.9, opacity: 0 }}
        animate={{ scale: 1, opacity: 1 }}
        className="z-10 group relative"
      >
        <Card className={cn(
          "p-4 text-center min-w-[240px] shadow-lg border-none transition-all hover:ring-2 hover:ring-primary/20",
          isRoot ? "bg-[#1A4A9A] text-white" : "bg-white border-primary/20 shadow-primary/5"
        )}>
          <div className="flex flex-col items-center">
            <Badge className={cn(
              "text-[8px] font-black uppercase tracking-[0.2em] mb-2 px-2 py-0.5",
              isRoot ? "bg-white/10 text-white" : "bg-primary/5 text-primary"
            )}>
              {isRoot ? 'Puncak Struktur' : 'Jabatan Struktural'}
            </Badge>
            <h4 className={cn("font-black text-xs uppercase tracking-tight leading-tight", !isRoot && "text-[#1A4A9A]")}>
              {node.namaJabatan}
            </h4>
            <div className={cn("mt-3 pt-3 border-t w-full", isRoot ? "border-white/10" : "border-primary/10")}>
              <p className={cn("text-[10px] font-bold italic", isRoot ? "text-white/90" : "text-primary/70")}>
                {node.namaPegawai}
              </p>
              <p className={cn("text-[8px] font-medium mt-1", isRoot ? "text-white/60" : "text-muted-foreground")}>
                NIP. {node.nip || '-'}
              </p>
            </div>
          </div>

          <div className="absolute top-2 right-2 flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
            <Button 
              variant="ghost" 
              size="icon" 
              className={cn("h-6 w-6 rounded-md", isRoot ? "hover:bg-white/10 text-white/50 hover:text-white" : "hover:bg-primary/5 text-muted-foreground hover:text-primary")}
              onClick={() => onEdit(node)}
            >
              <Pencil className="h-3 w-3" />
            </Button>
            <Button 
              variant="ghost" 
              size="icon" 
              className={cn("h-6 w-6 rounded-md", isRoot ? "hover:bg-red-500/20 text-white/50 hover:text-red-200" : "hover:bg-red-50 text-muted-foreground hover:text-red-500")}
              onClick={() => onDelete(node.id)}
            >
              <Trash2 className="h-3 w-3" />
            </Button>
          </div>

          <Button
            variant="ghost"
            size="sm"
            className={cn(
              "w-full h-7 mt-3 text-[8px] font-black uppercase tracking-widest opacity-0 group-hover:opacity-100 transition-all",
              isRoot ? "hover:bg-white/10 text-white/60 hover:text-white" : "hover:bg-primary/5 text-primary/60 hover:text-primary"
            )}
            onClick={() => onAddSub(node.id)}
          >
            + TAMBAH SUB-JABATAN
          </Button>
        </Card>
        
        {children.length > 0 && (
          <div className="absolute left-1/2 -bottom-10 w-0.5 h-10 bg-border -translate-x-1/2" />
        )}
      </motion.div>

      {children.length > 0 && (
        <div className="mt-10 flex justify-center items-start gap-8 relative pt-10">
          {children.length > 1 && (
            <div className="absolute top-0 left-[20%] right-[20%] h-0.5 bg-border rounded-full" />
          )}
          {children.map(child => (
            <div key={child.id} className="relative flex flex-col items-center">
              <div className="absolute -top-10 left-1/2 w-0.5 h-10 bg-border -translate-x-1/2" />
              <OrgNodeCard 
                node={child} 
                allNodes={allNodes} 
                onAddSub={onAddSub} 
                onEdit={onEdit} 
                onDelete={onDelete} 
              />
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [activeTab, setActiveTab] = useState<'struktur-organisasi' | 'tim-kerja' | 'skp' | 'skp-monitoring' | 'plt-plh' | 'arsip'>('tim-kerja');
  const [user, setUser] = useState<FirebaseUser | null>(null);
  const [isAuthLoading, setIsAuthLoading] = useState(true);
  const [records, setRecords] = useState<SKPRecord[]>([]);
  const [monitoringRecords, setMonitoringRecords] = useState<SKPMonitoringRecord[]>([]);
  const [pltRecords, setPltRecords] = useState<PLTRecord[]>([]);
  const [plhRecords, setPlhRecords] = useState<PLHRecord[]>([]);
  const [workTeams, setWorkTeams] = useState<WorkTeam[]>([]);
  const [unitInfoList, setUnitInfoList] = useState<UnitInfo[]>([]);
  const [orgNodes, setOrgNodes] = useState<OrgNode[]>([]);
  const [isEditingUnitSK, setIsEditingUnitSK] = useState(false);
  const [isAddOrgNodeDialogOpen, setIsAddOrgNodeDialogOpen] = useState(false);
  const [isEditOrgNodeDialogOpen, setIsEditOrgNodeDialogOpen] = useState(false);
  const [editingOrgNode, setEditingOrgNode] = useState<OrgNode | null>(null);
  const [parentOrgNodeId, setParentOrgNodeId] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [isImporting, setIsImporting] = useState(false);
  const [isImportingMonitoring, setIsImportingMonitoring] = useState(false);
  const [isImportingTeams, setIsImportingTeams] = useState(false);
  const [isImportingPltPlh, setIsImportingPltPlh] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const monitoringFileInputRef = useRef<HTMLInputElement>(null);
  const teamsFileInputRef = useRef<HTMLInputElement>(null);
  const pltPlhFileInputRef = useRef<HTMLInputElement>(null);

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
  const [isAddMonitoringDialogOpen, setIsAddMonitoringDialogOpen] = useState(false);
  const [isTakeDialogOpen, setIsTakeDialogOpen] = useState(false);
  const [isDetailDialogOpen, setIsDetailDialogOpen] = useState(false);
  const [isEditDialogOpen, setIsEditDialogOpen] = useState(false);
  const [isEditMonitoringDialogOpen, setIsEditMonitoringDialogOpen] = useState(false);
  const [selectedRecord, setSelectedRecord] = useState<SKPRecord | null>(null);
  const [detailRecord, setDetailRecord] = useState<SKPRecord | null>(null);
  const [editingRecord, setEditingRecord] = useState<SKPRecord | null>(null);
  const [editingMonitoringRecord, setEditingMonitoringRecord] = useState<SKPMonitoringRecord | null>(null);
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
  const [deletingInfo, setDeletingInfo] = useState<{ id: string, type: 'SKP' | 'SKP_MONITORING' | 'PLT' | 'PLH' | 'TIM' } | null>(null);
  const [monitoringSelectedEmployee, setMonitoringSelectedEmployee] = useState<string | null>(null);
  const [selectedUnitForTeams, setSelectedUnitForTeams] = useState<string | null>(null);
  const [selectedTeamIdForDetails, setSelectedTeamIdForDetails] = useState<string | null>(null);

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

    const qMonitoring = query(collection(db, 'skp_monitoring_records'), orderBy('createdAt', 'desc'));
    const unsubMonitoring = onSnapshot(qMonitoring, (snapshot) => {
      const docs = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as SKPMonitoringRecord[];
      setMonitoringRecords(docs);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'skp_monitoring_records');
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

    const qUnitInfo = query(collection(db, 'unit_info'));
    const unsubUnitInfo = onSnapshot(qUnitInfo, (snapshot) => {
      const docs = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as UnitInfo[];
      setUnitInfoList(docs);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'unit_info');
    });

    const qOrgStructure = query(collection(db, 'org_structure'), orderBy('order', 'asc'));
    const unsubOrgStructure = onSnapshot(qOrgStructure, (snapshot) => {
      const docs = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as OrgNode[];
      setOrgNodes(docs);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'org_structure');
    });

    return () => {
      unsubSkp();
      unsubMonitoring();
      unsubPlt();
      unsubPlh();
      unsubWorkTeams();
      unsubUnitInfo();
      unsubOrgStructure();
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

  const skpMonitoringStats = useMemo(() => {
    const total = monitoringRecords.length;
    const sangatBaik = monitoringRecords.filter(r => r.predikatKinerja === 'Sangat Baik').length;
    const baik = monitoringRecords.filter(r => r.predikatKinerja === 'Baik').length;
    const butuhPerbaikan = monitoringRecords.filter(r => r.predikatKinerja === 'Butuh Perbaikan').length;
    const kurus = monitoringRecords.filter(r => r.predikatKinerja === 'Kurang' || r.predikatKinerja === 'Sangat Kurang').length;
    return { total, sangatBaik, baik, butuhPerbaikan, kurang: kurus };
  }, [monitoringRecords]);

  const uniqueEmployeesForMonitoring = useMemo(() => {
    const employeeMap = new Map<string, { nip: string, namaPegawai: string, jabatan: string, unitKerja: string, latestPredikat?: string }>();
    
    // Sort records to get the latest info for each employee
    const sortedRecords = [...monitoringRecords].sort((a, b) => {
      if (b.tahun !== a.tahun) return b.tahun - a.tahun;
      const periodOrder = { 'Tahunan': 5, 'Triwulan IV': 4, 'Triwulan III': 3, 'Triwulan II': 2, 'Triwulan I': 1 };
      return (periodOrder[b.periode] || 0) - (periodOrder[a.periode] || 0);
    });

    sortedRecords.forEach(record => {
      const key = record.nip || record.namaPegawai;
      if (!employeeMap.has(key)) {
        employeeMap.set(key, {
          nip: record.nip,
          namaPegawai: record.namaPegawai,
          jabatan: record.jabatan,
          unitKerja: record.unitKerja,
          latestPredikat: record.predikatKinerja
        });
      }
    });

    const employees = Array.from(employeeMap.values());
    
    return employees.filter(emp => {
      const search = searchQuery.toLowerCase();
      return emp.namaPegawai.toLowerCase().includes(search) ||
             emp.nip.includes(search) ||
             emp.unitKerja.toLowerCase().includes(search);
    });
  }, [monitoringRecords, searchQuery]);

  const selectedEmployeeHistory = useMemo(() => {
    if (!monitoringSelectedEmployee) return [];
    return monitoringRecords
      .filter(r => (r.nip || r.namaPegawai) === monitoringSelectedEmployee)
      .sort((a, b) => {
        if (b.tahun !== a.tahun) return b.tahun - a.tahun;
        const periodOrder = { 'Tahunan': 5, 'Triwulan IV': 4, 'Triwulan III': 3, 'Triwulan II': 2, 'Triwulan I': 1 };
        return (periodOrder[b.periode] || 0) - (periodOrder[a.periode] || 0);
      });
  }, [monitoringRecords, monitoringSelectedEmployee]);

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
        authorUid: user.uid,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp(),
      };

      await addDoc(collection(db, 'skp_records'), newRecord);
      setIsAddDialogOpen(false);
      toast.success('Data SKP berhasil disimpan ke database');
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, 'skp_records');
    }
  };

  const handleAddMonitoringRecord = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!user) return;

    const formData = new FormData(e.currentTarget);
    
    try {
      const newRecord = {
        nip: formData.get('nip') as string,
        namaPegawai: formData.get('namaPegawai') as string,
        jabatan: formData.get('jabatan') as string,
        unitKerja: formData.get('unitKerja') as string,
        tahun: parseInt(formData.get('tahun') as string),
        periode: formData.get('periode') as SKPPeriode,
        ratingHasilKerja: formData.get('ratingHasilKerja') as string,
        ratingHasilPerilaku: formData.get('ratingHasilPerilaku') as string,
        predikatKinerja: formData.get('predikatKinerja') as string,
        authorUid: user.uid,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp(),
      };

      await addDoc(collection(db, 'skp_monitoring_records'), newRecord);
      setIsAddMonitoringDialogOpen(false);
      toast.success('Data Monitoring SKP berhasil disimpan');
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, 'skp_monitoring_records');
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
        tanggalSurat: formData.get('tanggalSurat') as string,
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
    
    try {
      const updatedData = {
        nip: formData.get('nip') as string,
        namaPegawai: formData.get('namaPegawai') as string,
        jabatan: formData.get('jabatan') as string,
        unitKerja: formData.get('unitKerja') as string,
        tahun: parseInt(formData.get('tahun') as string),
        periode: formData.get('periode') as SKPPeriode,
        jenisDokumen: selectedJenisDokumen,
        updatedAt: serverTimestamp(),
      };

      await updateDoc(doc(db, 'skp_records', editingRecord.id), updatedData);
      setIsEditDialogOpen(false);
      setEditingRecord(null);
      toast.success('Data SKP berhasil diperbarui');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `skp_records/${editingRecord.id}`);
    }
  };

  const handleUpdateMonitoringRecord = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!editingMonitoringRecord || !user) return;

    const formData = new FormData(e.currentTarget);
    
    try {
      const updatedData = {
        nip: formData.get('nip') as string,
        namaPegawai: formData.get('namaPegawai') as string,
        jabatan: formData.get('jabatan') as string,
        unitKerja: formData.get('unitKerja') as string,
        tahun: parseInt(formData.get('tahun') as string),
        periode: formData.get('periode') as SKPPeriode,
        ratingHasilKerja: formData.get('ratingHasilKerja') as string,
        ratingHasilPerilaku: formData.get('ratingHasilPerilaku') as string,
        predikatKinerja: formData.get('predikatKinerja') as string,
        updatedAt: serverTimestamp(),
      };

      await updateDoc(doc(db, 'skp_monitoring_records', editingMonitoringRecord.id), updatedData);
      setIsEditMonitoringDialogOpen(false);
      setEditingMonitoringRecord(null);
      toast.success('Data Monitoring SKP berhasil diperbarui');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, `skp_monitoring_records/${editingMonitoringRecord.id}`);
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
        tanggalSurat: formData.get('tanggalSurat') as string,
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

  const handleAddOrgNode = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!user) return;

    const formData = new FormData(e.currentTarget);
    const namaJabatan = formData.get('namaJabatan') as string;
    const namaPegawai = formData.get('namaPegawai') as string;
    const nip = formData.get('nip') as string;

    try {
      await addDoc(collection(db, 'org_structure'), {
        namaJabatan,
        namaPegawai,
        nip,
        parentId: parentOrgNodeId,
        order: orgNodes.length,
        authorUid: user.uid,
        updatedAt: serverTimestamp()
      });
      setIsAddOrgNodeDialogOpen(false);
      setParentOrgNodeId(null);
      toast.success('Jabatan struktural berhasil ditambahkan');
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'org_structure');
    }
  };

  const handleUpdateOrgNode = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!editingOrgNode || !user) return;

    const formData = new FormData(e.currentTarget);
    const namaJabatan = formData.get('namaJabatan') as string;
    const namaPegawai = formData.get('namaPegawai') as string;
    const nip = formData.get('nip') as string;

    try {
      await updateDoc(doc(db, 'org_structure', editingOrgNode.id), {
        namaJabatan,
        namaPegawai,
        nip,
        updatedAt: serverTimestamp()
      });
      setIsEditOrgNodeDialogOpen(false);
      setEditingOrgNode(null);
      toast.success('Jabatan struktural berhasil diperbarui');
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'org_structure');
    }
  };

  const handleDeleteOrgNode = async (id: string) => {
    if (!window.confirm('Hapus jabatan ini? Semua sub-jabatan di bawahnya harus dihapus manual atau dipindahkan.')) return;
    try {
      await deleteDoc(doc(db, 'org_structure', id));
      toast.success('Jabatan struktural berhasil dihapus');
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, 'org_structure');
    }
  };

  const handleUpdateUnitSK = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!selectedUnitForTeams || !user) return;

    const formData = new FormData(e.currentTarget);
    const nomorSK = formData.get('nomorSK') as string;
    const tanggalSK = formData.get('tanggalSK') as string;

    const existingInfo = unitInfoList.find(u => u.unitKerja === selectedUnitForTeams);

    try {
      if (existingInfo?.id) {
        await updateDoc(doc(db, 'unit_info', existingInfo.id), {
          nomorSK: nomorSK || '',
          tanggalSK: tanggalSK || '',
          authorUid: user.uid,
          updatedAt: serverTimestamp()
        });
      } else {
        await addDoc(collection(db, 'unit_info'), {
          unitKerja: selectedUnitForTeams,
          nomorSK: nomorSK || '',
          tanggalSK: tanggalSK || '',
          authorUid: user.uid,
          updatedAt: serverTimestamp()
        });
      }
      setIsEditingUnitSK(false);
      toast.success('Informasi SK Unit Kerja berhasil diperbarui');
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'unit_info');
    }
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
    const collectionName = 
      type === 'SKP' ? 'skp_records' : 
      type === 'SKP_MONITORING' ? 'skp_monitoring_records' :
      type === 'PLT' ? 'plt_records' : 
      type === 'PLH' ? 'plh_records' : 'work_teams';
    
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

  const triggerFileInput = () => {
    fileInputRef.current?.click();
  };

  const triggerMonitoringFileInput = () => {
    monitoringFileInputRef.current?.click();
  };

  const handleImportExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Check file extension
    const fileExt = file.name.split('.').pop()?.toLowerCase();
    if (fileExt !== 'xlsx' && fileExt !== 'xls') {
      toast.error('Format file harus .xlsx atau .xls');
      return;
    }

    setIsImporting(true);
    const reader = new FileReader();

    reader.onload = async (evt) => {
      try {
        const dataBuffer = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(dataBuffer, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        if (data.length === 0) {
          toast.error('File Excel kosong');
          setIsImporting(false);
          return;
        }

        let importCount = 0;
        let skipCount = 0;
        
        for (let i = 0; i < data.length; i++) {
          const row = data[i] as any;
          
          try {
            // Get raw values with multiple possible aliases
            const rawNama = row['Nama Pegawai'] || row['namaPegawai'] || row['Nama'] || row['nama'] || '';
            const rawNip = row['NIP'] || row['nip'] || '';
            const rawJabatan = row['Jabatan'] || row['jabatan'] || '';
            const rawUnit = row['Unit Kerja'] || row['unitKerja'] || '';
            const rawTahun = row['Tahun'] || row['tahun'] || row['thn'] || new Date().getFullYear();
            const rawPeriode = row['Periode'] || row['periode'] || 'Tahunan';
            const rawJenisDok = row['Jenis Dokumen'] || row['jenisDokumen'] || row['dokumen'] || '';
            const rawStatus = row['Status'] || row['status'] || 'Sudah Diambil';
            
            // Rating attributes
            const rawRatingHasil = row['Rating Hasil Kerja'] || row['ratingHasilKerja'] || row['Rating Kerja'] || row['hasilKerja'] || null;
            const rawRatingPerilaku = row['Rating Hasil Perilaku'] || row['ratingHasilPerilaku'] || row['Rating Perilaku'] || row['perilaku'] || null;
            const rawPredikat = row['Predikat Kinerja'] || row['predikatKinerja'] || row['Predikat'] || row['predikat'] || null;

            // Map fields with type conversion
            const newRecord = {
              nip: String(rawNip).trim(),
              namaPegawai: String(rawNama).trim(),
              jabatan: String(rawJabatan).trim(),
              unitKerja: String(rawUnit).trim(),
              tahun: isNaN(Number(rawTahun)) ? new Date().getFullYear() : Number(rawTahun),
              periode: (PERIODE_LIST.includes(rawPeriode as SKPPeriode) ? rawPeriode : 'Tahunan') as SKPPeriode,
              jenisDokumen: Array.isArray(rawJenisDok) 
                ? rawJenisDok 
                : String(rawJenisDok).split(/[,,|]/).map(s => s.trim()).filter(Boolean),
              status: (['Sudah Diambil', 'Belum Diambil'].includes(rawStatus as SKPStatus) ? rawStatus : 'Sudah Diambil') as SKPStatus,
              ratingHasilKerja: rawRatingHasil ? String(rawRatingHasil).trim() : null,
              ratingHasilPerilaku: rawRatingPerilaku ? String(rawRatingPerilaku).trim() : null,
              predikatKinerja: rawPredikat ? String(rawPredikat).trim() : null,
              authorUid: user.uid,
              createdAt: serverTimestamp(),
              updatedAt: serverTimestamp(),
            };

            // Validation: Minimal requirement is Name
            if (!newRecord.namaPegawai) {
              console.warn(`Row ${i + 1} skipped: Nama Pegawai kosong`);
              skipCount++;
              continue;
            }

            await addDoc(collection(db, 'skp_records'), newRecord);
            importCount++;
          } catch (rowError) {
            console.error(`Error processing row ${i + 1}:`, rowError);
            handleFirestoreError(rowError, OperationType.CREATE, 'skp_records_batch_import');
          }
        }

        if (importCount > 0) {
          toast.success(`${importCount} data SKP berhasil diimpor.${skipCount > 0 ? ` (${skipCount} baris dilewati because data tidak lengkap)` : ''}`);
        } else {
          toast.error('Tidak ada data valid yang berhasil diimpor. Pastikan kolom "Nama Pegawai" terisi.');
        }
      } catch (error) {
        console.error('Error importing Excel:', error);
        toast.error('Gagal mengimpor file Excel. Pastikan file tidak diproteksi dan format kolom mendekati standar export.');
      } finally {
        setIsImporting(false);
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    };

    reader.readAsArrayBuffer(file);
  };
  
  const triggerPltPlhFileInput = () => {
    pltPlhFileInputRef.current?.click();
  };
  
  const triggerTeamsFileInput = () => {
    teamsFileInputRef.current?.click();
  };

  const handleExportTeams = () => {
    if (workTeams.length === 0) {
      toast.error('Tidak ada data tim untuk diekspor');
      return;
    }

    const workbook = XLSX.utils.book_new();
    
    // Group teams by unitKerja
    const groupedExportData: Record<string, any[]> = {};
    
    workTeams.forEach(team => {
      const unit = team.unitKerja || 'Lainnya';
      const unitInfo = unitInfoList.find(u => u.unitKerja === unit);

      const data = {
        'Nama Tim': team.namaTim,
        'Unit Kerja': team.unitKerja,
        'Tahun': team.tahun,
        'Status Tim': team.status,
        'Jenis': team.jenis,
        'Nama Ketua': team.ketua.nama,
        'Jabatan Ketua': team.ketua.jabatan,
        'Status Ketua': team.ketua.status,
        'Nomor SK Unit': unitInfo?.nomorSK || '-',
        'Tanggal SK Unit': unitInfo?.tanggalSK || '-',
        'Jumlah Anggota': team.anggota.length,
        'Daftar Anggota': team.anggota.map(a => `${a.nama} (${a.jabatan})`).join('; '),
        'Bagian': team.parentId ? (workTeams.find(t => t.id === team.parentId)?.namaTim || '-') : '-'
      };
      
      if (!groupedExportData[unit]) {
        groupedExportData[unit] = [];
      }
      groupedExportData[unit].push(data);
    });

    // Create a sheet for each Unit Kerja
    Object.entries(groupedExportData).forEach(([unit, data]) => {
      // Excel sheet name limits: max 31 chars, no special chars \ / ? * [ ] :
      const safeUnitName = unit.substring(0, 31).replace(/[\\/?*[\]:]/g, '_') || 'Sheet';
      const worksheet = XLSX.utils.json_to_sheet(data);
      XLSX.utils.book_append_sheet(workbook, worksheet, safeUnitName);
    });

    XLSX.writeFile(workbook, `Tim_Kerja_${format(new Date(), 'yyyyMMdd')}.xlsx`);
  };

  const handleImportTeamsExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const fileExt = file.name.split('.').pop()?.toLowerCase();
    if (fileExt !== 'xlsx' && fileExt !== 'xls') {
      toast.error('Format file harus .xlsx atau .xls');
      return;
    }

    setIsImportingTeams(true);
    const reader = new FileReader();

    reader.onload = async (evt) => {
      try {
        const dataBuffer = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(dataBuffer, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        if (data.length === 0) {
          toast.error('File Excel kosong');
          setIsImportingTeams(false);
          return;
        }

        let importCount = 0;
        let skipCount = 0;
        
        for (let i = 0; i < data.length; i++) {
          const row = data[i] as any;
          
          try {
            const namaTim = row['Nama Tim'] || row['namaTim'] || '';
            const unitKerja = row['Unit Kerja'] || row['unitKerja'] || '';
            const namaKetua = row['Nama Ketua'] || row['namaKetua'] || '';
            const jabatanKetua = row['Jabatan Ketua'] || row['jabatanKetua'] || '';
            const statusKetua = row['Status Ketua'] || row['statusKetua'] || 'PNS';
            const tahun = row['Tahun'] || row['tahun'] || new Date().getFullYear();
            const jenis = row['Jenis'] || row['jenis'] || 'Tim Kerja';
            const statusTeam = row['Status Tim'] || row['statusTim'] || 'Aktif';

            if (!namaTim || !unitKerja || !namaKetua || !jabatanKetua) {
              skipCount++;
              continue;
            }

            const newTeam = {
              namaTim: String(namaTim).trim(),
              unitKerja: String(unitKerja).trim(),
              tahun: isNaN(Number(tahun)) ? new Date().getFullYear() : Number(tahun),
              status: (statusTeam === 'Aktif' || statusTeam === 'Non-Aktif' ? statusTeam : 'Aktif') as 'Aktif' | 'Non-Aktif',
              jenis: (jenis === 'Tim Kerja' || jenis === 'Bagian' ? jenis : 'Tim Kerja') as 'Tim Kerja' | 'Bagian',
              ketua: {
                nama: String(namaKetua).trim(),
                jabatan: String(jabatanKetua).trim(),
                status: String(statusKetua).trim()
              },
              anggota: [],
              authorUid: user!.uid,
              createdAt: serverTimestamp(),
              updatedAt: serverTimestamp(),
            };

            await addDoc(collection(db, 'work_teams'), newTeam);
            importCount++;
          } catch (err) {
            console.error('Error importing team row:', err);
            skipCount++;
          }
        }

        toast.success(`Berhasil mengimpor ${importCount} tim kerja. Skip: ${skipCount}`);
      } catch (error) {
        console.error('Import error:', error);
        toast.error('Gagal membaca file Excel');
      } finally {
        setIsImportingTeams(false);
        if (teamsFileInputRef.current) teamsFileInputRef.current.value = '';
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const handleImportMonitoringExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const fileExt = file.name.split('.').pop()?.toLowerCase();
    if (fileExt !== 'xlsx' && fileExt !== 'xls') {
      toast.error('Format file harus .xlsx atau .xls');
      return;
    }

    setIsImportingMonitoring(true);
    const reader = new FileReader();

    reader.onload = async (evt) => {
      try {
        const dataBuffer = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(dataBuffer, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        if (data.length === 0) {
          toast.error('File Excel kosong');
          setIsImportingMonitoring(false);
          return;
        }

        let importCount = 0;
        let skipCount = 0;
        
        for (let i = 0; i < data.length; i++) {
          const row = data[i] as any;
          
          try {
            const rawNama = row['Nama Pegawai'] || row['namaPegawai'] || row['Nama'] || row['nama'] || '';
            const rawNip = row['NIP'] || row['nip'] || '';
            const rawJabatan = row['Jabatan'] || row['jabatan'] || '';
            const rawUnit = row['Unit Kerja'] || row['unitKerja'] || '';
            const rawTahun = row['Tahun'] || row['tahun'] || new Date().getFullYear();
            const rawPeriode = row['Periode'] || row['periode'] || 'Tahunan';
            const rawRatingHasil = row['Rating Hasil Kerja'] || row['ratingHasilKerja'] || null;
            const rawRatingPerilaku = row['Rating Hasil Perilaku'] || row['ratingHasilPerilaku'] || null;
            const rawPredikat = row['Predikat Kinerja'] || row['predikatKinerja'] || null;

            const newRecord = {
              nip: String(rawNip).trim(),
              namaPegawai: String(rawNama).trim(),
              jabatan: String(rawJabatan).trim(),
              unitKerja: String(rawUnit).trim(),
              tahun: isNaN(Number(rawTahun)) ? new Date().getFullYear() : Number(rawTahun),
              periode: (PERIODE_LIST.includes(rawPeriode as SKPPeriode) ? rawPeriode : 'Tahunan') as SKPPeriode,
              ratingHasilKerja: rawRatingHasil ? String(rawRatingHasil).trim() : null,
              ratingHasilPerilaku: rawRatingPerilaku ? String(rawRatingPerilaku).trim() : null,
              predikatKinerja: rawPredikat ? String(rawPredikat).trim() : null,
              authorUid: user.uid,
              createdAt: serverTimestamp(),
              updatedAt: serverTimestamp(),
            };

            if (!newRecord.namaPegawai) {
              skipCount++;
              continue;
            }

            await addDoc(collection(db, 'skp_monitoring_records'), newRecord);
            importCount++;
          } catch (err) {
            console.error('Error importing row:', err);
            skipCount++;
          }
        }

        toast.success(`Berhasil mengimpor ${importCount} data monitoring. Skip: ${skipCount}`);
      } catch (error) {
        console.error('Import error:', error);
        toast.error('Gagal membaca file Excel');
      } finally {
        setIsImportingMonitoring(false);
        if (monitoringFileInputRef.current) monitoringFileInputRef.current.value = '';
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const handleImportPltPlhExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!user) {
      toast.error('Silakan login terlebih dahulu');
      return;
    }

    const fileExt = file.name.split('.').pop()?.toLowerCase();
    if (fileExt !== 'xlsx' && fileExt !== 'xls') {
      toast.error('Format file harus .xlsx atau .xls');
      return;
    }

    setIsImportingPltPlh(true);
    const reader = new FileReader();

    reader.onload = async (evt) => {
      try {
        const dataBuffer = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(dataBuffer, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        if (data.length === 0) {
          toast.error('File Excel kosong');
          setIsImportingPltPlh(false);
          return;
        }

        let pltCount = 0;
        let plhCount = 0;
        let skipCount = 0;
        
        for (let i = 0; i < data.length; i++) {
          const row = data[i] as any;
          
          try {
            const rawNama = row['Nama Pegawai'] || row['namaPegawai'] || row['Nama'] || '';
            const rawNip = row['NIP'] || row['nip'] || '';
            const rawTipe = String(row['Tipe'] || row['tipe'] || row['Jenis'] || '').toUpperCase();
            const isPLT = rawTipe.includes('PLT');
            const isPLH = rawTipe.includes('PLH');

            if (!isPLT && !isPLH) {
              console.warn(`Row ${i + 1} skipped: Tipe tidak jelas (PLT/PLH)`);
              skipCount++;
              continue;
            }

            const rawJabatanAsli = row['Jabatan Asli'] || row['jabatanAsli'] || '';
            const rawJabatanTugas = row['Jabatan Tugas'] || row['jabatanTugas'] || row['Jabatan PLT'] || row['Jabatan PLH'] || '';
            const rawUnitAsli = row['Unit Kerja Asli'] || row['unitKerja'] || '';
            const rawUnitTugas = row['Unit Kerja Tugas'] || row['unitKerjaTugas'] || '';
            const rawNoSk = row['No SK'] || row['noSk'] || '';
            const rawTglMulai = row['Tanggal Mulai'] || row['tglMulai'] || '';
            const rawTglSelesai = row['Tanggal Selesai'] || row['tglSelesai'] || '';
            const rawStatus = row['Status'] || row['status'] || 'Aktif';

            if (!rawNama) {
              skipCount++;
              continue;
            }

            if (isPLT) {
              const newPlt = {
                nip: String(rawNip).trim(),
                namaPegawai: String(rawNama).trim(),
                jabatanAsli: String(rawJabatanAsli).trim(),
                jabatanPlt: String(rawJabatanTugas).trim(),
                unitKerja: String(rawUnitAsli).trim(),
                unitKerjaTugas: String(rawUnitTugas).trim(),
                noSk: String(rawNoSk).trim(),
                tglMulai: String(rawTglMulai).trim(),
                tglSelesai: String(rawTglSelesai).trim(),
                status: (rawStatus === 'Selesai' ? 'Selesai' : 'Aktif') as PLTStatus,
                authorUid: user.uid,
                createdAt: serverTimestamp(),
                updatedAt: serverTimestamp(),
              };
              await addDoc(collection(db, 'plt_records'), newPlt);
              pltCount++;
            } else {
              const newPlh = {
                nip: String(rawNip).trim(),
                namaPegawai: String(rawNama).trim(),
                jabatanAsli: String(rawJabatanAsli).trim(),
                jabatanPlh: String(rawJabatanTugas).trim(),
                unitKerja: String(rawUnitAsli).trim(),
                unitKerjaTugas: String(rawUnitTugas).trim(),
                noSk: String(rawNoSk).trim(),
                tglMulai: String(rawTglMulai).trim(),
                tglSelesai: String(rawTglSelesai).trim(),
                status: (rawStatus === 'Selesai' ? 'Selesai' : 'Aktif') as PLTStatus,
                authorUid: user.uid,
                createdAt: serverTimestamp(),
                updatedAt: serverTimestamp(),
              };
              await addDoc(collection(db, 'plh_records'), newPlh);
              plhCount++;
            }
          } catch (rowError) {
            console.error(`Error processing row ${i + 1}:`, rowError);
          }
        }

        toast.success(`Berhasil mengimpor ${pltCount} PLT dan ${plhCount} PLH.${skipCount > 0 ? ` (${skipCount} baris dilewati)` : ''}`);
      } catch (error) {
        console.error('Error importing Excel:', error);
        toast.error('Gagal mengimpor file Excel PLT/PLH.');
      } finally {
        setIsImportingPltPlh(false);
        if (pltPlhFileInputRef.current) pltPlhFileInputRef.current.value = '';
      }
    };

    reader.readAsArrayBuffer(file);
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
        'Tanggal Surat': record.tanggalSurat ? format(new Date(record.tanggalSurat), 'dd MMMM yyyy', { locale: id }) : '-',
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
        'Tanggal Surat': record.tanggalSurat ? format(new Date(record.tanggalSurat), 'dd MMMM yyyy', { locale: id }) : '-',
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
    <Card key={team.id} className="overflow-hidden border-none shadow-sm hover:shadow-md transition-all group cursor-pointer" onClick={() => setSelectedTeamIdForDetails(team.id)}>
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
          <div className="h-8 w-8 rounded-full flex items-center justify-center text-primary/20 group-hover:bg-primary group-hover:text-white transition-all">
            <ArrowRight className="h-4 w-4" />
          </div>
        </div>
      </CardHeader>
      <CardContent className="pt-4">
        <div className="flex items-center justify-between text-[10px] text-muted-foreground font-bold uppercase tracking-widest">
          <div className="flex items-center gap-1.5 font-black text-[#1A4A9A] line-clamp-1">
            <UserCheck className="h-3 w-3 shrink-0" />
            <span className="truncate">{team.ketua.nama}</span>
          </div>
          <div className="flex items-center gap-1.5">
            <Users className="h-3 w-3" />
            {team.anggota.length} Anggota
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
              variant={activeTab === 'struktur-organisasi' ? 'default' : 'ghost'} 
              className={cn(
                "justify-start gap-3 h-10 px-4 transition-all", 
                activeTab === 'struktur-organisasi' ? "shadow-sm" : "text-muted-foreground"
              )}
              onClick={() => setActiveTab('struktur-organisasi')}
            >
              <Network className="h-4 w-4" />
              Struktur Organisasi
            </Button>
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
              <FileCheck className="h-4 w-4" />
              Berkas SKP Pegawai
            </Button>
            <Button 
              variant={activeTab === 'skp-monitoring' ? 'default' : 'ghost'} 
              className={cn(
                "justify-start gap-3 h-10 px-4 transition-all", 
                activeTab === 'skp-monitoring' ? "shadow-sm" : "text-muted-foreground"
              )}
              onClick={() => {
                setActiveTab('skp-monitoring');
                setMonitoringSelectedEmployee(null);
              }}
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
            <Button 
              variant={activeTab === 'arsip' ? 'default' : 'ghost'} 
              className={cn(
                "justify-start gap-3 h-10 px-4 transition-all", 
                activeTab === 'arsip' ? "shadow-sm" : "text-muted-foreground"
              )}
              onClick={() => setActiveTab('arsip')}
            >
              <Archive className="h-4 w-4" />
              Arsip Dokumen
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
                {activeTab === 'struktur-organisasi' ? 'Struktur Organisasi' : activeTab === 'tim-kerja' ? 'Monitoring Tim Kerja' : activeTab === 'skp' ? 'Monitoring Berkas SKP Pegawai' : activeTab === 'skp-monitoring' ? 'Monitoring SKP Pegawai (Rating)' : activeTab === 'plt-plh' ? 'Monitoring PLT & PLH' : 'Arsip Dokumen'}
              </h2>
            </div>

            <div className="flex items-center gap-3">
              {user && activeTab === 'tim-kerja' && (
                <>
                  <input
                    type="file"
                    ref={teamsFileInputRef}
                    onChange={handleImportTeamsExcel}
                    accept=".xlsx, .xls"
                    className="hidden"
                  />
                  <Button 
                    variant="outline" 
                    size="sm" 
                    className="hidden sm:flex gap-2 h-9"
                    onClick={triggerTeamsFileInput}
                    disabled={isImportingTeams}
                  >
                    {isImportingTeams ? (
                      <span className="flex items-center gap-2">
                        <span className="h-3 w-3 border-2 border-primary border-t-transparent rounded-full animate-spin" />
                        Importing...
                      </span>
                    ) : (
                      <>
                        <Upload className="h-4 w-4" />
                        Import Excel
                      </>
                    )}
                  </Button>
                  <Button 
                    variant="outline" 
                    size="sm" 
                    className="hidden sm:flex gap-2 h-9"
                    onClick={handleExportTeams}
                  >
                    <Download className="h-4 w-4" />
                    Export Excel
                  </Button>
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
                          <Select name="unitKerja" defaultValue={selectedUnitForTeams || undefined} required>
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
              </>
              )}

              {user && activeTab === 'skp' && (
                <>
                  <input
                    type="file"
                    ref={fileInputRef}
                    onChange={handleImportExcel}
                    accept=".xlsx, .xls"
                    className="hidden"
                  />
                  <Button 
                    variant="outline" 
                    size="sm" 
                    className="hidden sm:flex gap-2 h-9"
                    onClick={triggerFileInput}
                    disabled={isImporting}
                  >
                    {isImporting ? (
                      <span className="flex items-center gap-2">
                        <span className="h-3 w-3 border-2 border-primary border-t-transparent rounded-full animate-spin" />
                        Importing...
                      </span>
                    ) : (
                      <>
                        <Upload className="h-4 w-4" />
                        Import Excel
                      </>
                    )}
                  </Button>
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
                  }}>
                    <DialogTrigger 
                      nativeButton={true}
                      render={<Button size="sm" className="gap-2 h-9 shadow-sm" />}
                    >
                      <Plus className="h-4 w-4" />
                      Tambah Data Berkas
                    </DialogTrigger>
                    <DialogContent className="sm:max-w-[550px] max-h-[90vh] overflow-y-auto">
                      {isAddDialogOpen && (
                        <form onSubmit={handleAddRecord}>
                          <DialogHeader>
                            <DialogTitle>Tambah Berkas SKP Baru</DialogTitle>
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
                          </div>
                          <DialogFooter className="gap-2 sm:gap-0 sticky bottom-0 bg-background pt-2 border-t mt-2">
                            <DialogClose 
                              nativeButton={true}
                              render={<Button type="button" variant="outline">Batal</Button>} 
                            />
                            <Button type="submit" className="bg-primary hover:bg-primary/90">
                              Simpan Berkas
                            </Button>
                          </DialogFooter>
                      </form>
                    )}
                    </DialogContent>
                  </Dialog>
                </>
              )}

              {user && activeTab === 'skp-monitoring' && (
                <>
                  <input
                    type="file"
                    ref={monitoringFileInputRef}
                    onChange={handleImportMonitoringExcel}
                    accept=".xlsx, .xls"
                    className="hidden"
                  />
                  <Button 
                    variant="outline" 
                    size="sm" 
                    className="hidden sm:flex gap-2 h-9"
                    onClick={triggerMonitoringFileInput}
                    disabled={isImportingMonitoring}
                  >
                    {isImportingMonitoring ? (
                      <span className="flex items-center gap-2">
                        <span className="h-3 w-3 border-2 border-primary border-t-transparent rounded-full animate-spin" />
                        Importing...
                      </span>
                    ) : (
                      <>
                        <Upload className="h-4 w-4" />
                        Import Excel
                      </>
                    )}
                  </Button>
                  <Button 
                    variant="outline" 
                    size="sm" 
                    className="hidden sm:flex gap-2 h-9"
                    onClick={() => {
                      if (monitoringRecords.length === 0) {
                        toast.error('Tidak ada data untuk diekspor');
                        return;
                      }
                      const exportData = monitoringRecords.map(r => ({
                        'NIP': r.nip,
                        'Nama Pegawai': r.namaPegawai,
                        'Jabatan': r.jabatan,
                        'Unit Kerja': r.unitKerja,
                        'Tahun': r.tahun,
                        'Periode': r.periode,
                        'Rating Hasil Kerja': r.ratingHasilKerja || '-',
                        'Rating Hasil Perilaku': r.ratingHasilPerilaku || '-',
                        'Predikat Kinerja': r.predikatKinerja || '-',
                      }));
                      const worksheet = XLSX.utils.json_to_sheet(exportData);
                      const workbook = XLSX.utils.book_new();
                      XLSX.utils.book_append_sheet(workbook, worksheet, 'Data Monitoring SKP');
                      XLSX.writeFile(workbook, `Monitoring_SKP_${format(new Date(), 'yyyyMMdd')}.xlsx`);
                    }}
                  >
                    <Download className="h-4 w-4" />
                    Export Monitoring
                  </Button>
                  <Dialog open={isAddMonitoringDialogOpen} onOpenChange={setIsAddMonitoringDialogOpen}>
                    <DialogTrigger 
                      nativeButton={true}
                      render={<Button size="sm" className="gap-2 h-9 shadow-sm" />}
                    >
                      <Plus className="h-4 w-4" />
                      Tambah Nilai SKP
                    </DialogTrigger>
                    <DialogContent className="sm:max-w-[550px] max-h-[90vh] overflow-y-auto">
                      <form onSubmit={handleAddMonitoringRecord}>
                        <DialogHeader>
                          <DialogTitle>Tambah Nilai SKP Baru</DialogTitle>
                          <DialogDescription>
                            Masukkan penilaian rating dan predikat kinerja pegawai.
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
                                    {UNIT_KERJA_LIST.map(v => <SelectItem key={v} value={v}>{v}</SelectItem>)}
                                  </SelectContent>
                                </Select>
                              </div>
                            </div>

                            <div className="grid grid-cols-2 gap-4">
                              <div className="grid gap-2">
                                <Label htmlFor="tahun">Tahun</Label>
                                <Input name="tahun" type="number" defaultValue={new Date().getFullYear()} required />
                              </div>
                              <div className="grid gap-2">
                                <Label htmlFor="periode">Periode</Label>
                                <Select name="periode" defaultValue="Tahunan" required>
                                  <SelectTrigger>
                                    <SelectValue placeholder="Pilih Periode" />
                                  </SelectTrigger>
                                  <SelectContent>
                                    {PERIODE_LIST.map(v => <SelectItem key={v} value={v}>{v}</SelectItem>)}
                                  </SelectContent>
                                </Select>
                              </div>
                            </div>

                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                              <div className="grid gap-2">
                                <Label>Rating Hasil Kerja</Label>
                                <Select name="ratingHasilKerja">
                                  <SelectTrigger>
                                    <SelectValue placeholder="Pilih Rating" />
                                  </SelectTrigger>
                                  <SelectContent>
                                    {RATING_HASIL_KERJA_LIST.map(v => <SelectItem key={v} value={v}>{v}</SelectItem>)}
                                  </SelectContent>
                                </Select>
                              </div>
                              <div className="grid gap-2">
                                <Label>Rating Perilaku</Label>
                                <Select name="ratingHasilPerilaku">
                                  <SelectTrigger>
                                    <SelectValue placeholder="Pilih Rating" />
                                  </SelectTrigger>
                                  <SelectContent>
                                    {RATING_PERILAKU_LIST.map(v => <SelectItem key={v} value={v}>{v}</SelectItem>)}
                                  </SelectContent>
                                </Select>
                              </div>
                            </div>

                            <div className="grid gap-2">
                              <Label>Predikat Kinerja</Label>
                              <Select name="predikatKinerja">
                                <SelectTrigger>
                                  <SelectValue placeholder="Pilih Predikat" />
                                </SelectTrigger>
                                <SelectContent>
                                  {PREDIKAT_KINERJA_LIST.map(v => <SelectItem key={v} value={v}>{v}</SelectItem>)}
                                </SelectContent>
                              </Select>
                            </div>
                        </div>
                        <DialogFooter className="sticky bottom-0 bg-background pt-2 border-t mt-2">
                          <DialogClose nativeButton={true} render={<Button type="button" variant="outline">Batal</Button>} />
                          <Button type="submit">Simpan Nilai</Button>
                        </DialogFooter>
                      </form>
                    </DialogContent>
                  </Dialog>

                  <Dialog open={isEditMonitoringDialogOpen} onOpenChange={setIsEditMonitoringDialogOpen}>
                    <DialogContent className="sm:max-w-[550px] max-h-[90vh] overflow-y-auto">
                      {editingMonitoringRecord && (
                        <form onSubmit={handleUpdateMonitoringRecord}>
                          <DialogHeader>
                            <DialogTitle>Edit Nilai SKP</DialogTitle>
                            <DialogDescription>
                              Perbarui penilaian rating dan predikat kinerja pegawai.
                            </DialogDescription>
                          </DialogHeader>
                          <div className="grid gap-4 py-4 max-h-[65vh] overflow-y-auto pr-2 custom-scrollbar">
                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                <div className="grid gap-2">
                                  <Label htmlFor="editMonitoringNamaPegawai">Nama Lengkap</Label>
                                  <Input 
                                    id="editMonitoringNamaPegawai" 
                                    name="namaPegawai" 
                                    defaultValue={editingMonitoringRecord.namaPegawai} 
                                    required 
                                  />
                                </div>
                                <div className="grid gap-2">
                                  <Label htmlFor="editMonitoringNip">NIP</Label>
                                  <Input 
                                    id="editMonitoringNip" 
                                    name="nip" 
                                    defaultValue={editingMonitoringRecord.nip} 
                                    required 
                                  />
                                </div>
                              </div>
                              
                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                <div className="grid gap-2">
                                  <Label htmlFor="editMonitoringJabatan">Jabatan</Label>
                                  <Input 
                                    id="editMonitoringJabatan" 
                                    name="jabatan" 
                                    defaultValue={editingMonitoringRecord.jabatan} 
                                    required 
                                  />
                                </div>
                                <div className="grid gap-2">
                                  <Label htmlFor="editMonitoringUnitKerja">Unit Kerja</Label>
                                  <Select name="unitKerja" defaultValue={editingMonitoringRecord.unitKerja}>
                                    <SelectTrigger>
                                      <SelectValue placeholder="Pilih Unit Kerja" />
                                    </SelectTrigger>
                                    <SelectContent>
                                      {UNIT_KERJA_LIST.map(v => <SelectItem key={v} value={v}>{v}</SelectItem>)}
                                    </SelectContent>
                                  </Select>
                                </div>
                              </div>

                              <div className="grid grid-cols-2 gap-4">
                                <div className="grid gap-2">
                                  <Label htmlFor="editMonitoringTahun">Tahun</Label>
                                  <Input name="tahun" type="number" defaultValue={editingMonitoringRecord.tahun} required />
                                </div>
                                <div className="grid gap-2">
                                  <Label htmlFor="editMonitoringPeriode">Periode</Label>
                                  <Select name="periode" defaultValue={editingMonitoringRecord.periode}>
                                    <SelectTrigger>
                                      <SelectValue placeholder="Pilih Periode" />
                                    </SelectTrigger>
                                    <SelectContent>
                                      {PERIODE_LIST.map(v => <SelectItem key={v} value={v}>{v}</SelectItem>)}
                                    </SelectContent>
                                  </Select>
                                </div>
                              </div>

                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                <div className="grid gap-2">
                                  <Label>Rating Hasil Kerja</Label>
                                  <Select name="ratingHasilKerja" defaultValue={editingMonitoringRecord.ratingHasilKerja}>
                                    <SelectTrigger>
                                      <SelectValue placeholder="Pilih Rating" />
                                    </SelectTrigger>
                                    <SelectContent>
                                      {RATING_HASIL_KERJA_LIST.map(v => <SelectItem key={v} value={v}>{v}</SelectItem>)}
                                    </SelectContent>
                                  </Select>
                                </div>
                                <div className="grid gap-2">
                                  <Label>Rating Perilaku</Label>
                                  <Select name="ratingHasilPerilaku" defaultValue={editingMonitoringRecord.ratingHasilPerilaku}>
                                    <SelectTrigger>
                                      <SelectValue placeholder="Pilih Rating" />
                                    </SelectTrigger>
                                    <SelectContent>
                                      {RATING_PERILAKU_LIST.map(v => <SelectItem key={v} value={v}>{v}</SelectItem>)}
                                    </SelectContent>
                                  </Select>
                                </div>
                              </div>

                              <div className="grid gap-2">
                                <Label>Predikat Kinerja</Label>
                                <Select name="predikatKinerja" defaultValue={editingMonitoringRecord.predikatKinerja}>
                                  <SelectTrigger>
                                    <SelectValue placeholder="Pilih Predikat" />
                                  </SelectTrigger>
                                  <SelectContent>
                                    {PREDIKAT_KINERJA_LIST.map(v => <SelectItem key={v} value={v}>{v}</SelectItem>)}
                                  </SelectContent>
                                </Select>
                              </div>
                          </div>
                          <DialogFooter className="sticky bottom-0 bg-background pt-2 border-t mt-2">
                            <DialogClose nativeButton={true} render={<Button type="button" variant="outline">Batal</Button>} />
                            <Button type="submit">Update Nilai</Button>
                          </DialogFooter>
                        </form>
                      )}
                    </DialogContent>
                  </Dialog>
                </>
              )}

              {user && activeTab === 'plt-plh' && (
                <>
                  <input
                    type="file"
                    ref={pltPlhFileInputRef}
                    onChange={handleImportPltPlhExcel}
                    accept=".xlsx, .xls"
                    className="hidden"
                  />
                  <Button 
                    variant="outline" 
                    size="sm" 
                    className="hidden sm:flex gap-2 h-9"
                    onClick={triggerPltPlhFileInput}
                    disabled={isImportingPltPlh}
                  >
                    {isImportingPltPlh ? (
                      <span className="flex items-center gap-2">
                        <span className="h-3 w-3 border-2 border-primary border-t-transparent rounded-full animate-spin" />
                        Importing...
                      </span>
                    ) : (
                      <>
                        <Upload className="h-4 w-4" />
                        Import Excel
                      </>
                    )}
                  </Button>
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

                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div className="grid gap-2">
                              <Label htmlFor="noSk">Nomor SK</Label>
                              <Input id="noSk" name="noSk" placeholder="Nomor Surat Keputusan" required />
                            </div>
                            <div className="grid gap-2">
                              <Label htmlFor="tanggalSurat">Tanggal Surat</Label>
                              <Input id="tanggalSurat" name="tanggalSurat" type="date" required />
                            </div>
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
        {activeTab === 'struktur-organisasi' ? (
          <div className="space-y-8">
            <motion.div 
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-white/80 backdrop-blur-md p-8 rounded-3xl shadow-sm border border-primary/10"
            >
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 mb-8">
                <div>
                  <h2 className="text-3xl font-black text-[#1A4A9A] tracking-tight">Struktur Organisasi</h2>
                  <p className="text-muted-foreground font-medium mt-1 uppercase tracking-widest text-[10px]">Hierarki Struktural Formal Satuan Kerja</p>
                </div>
                <Button 
                  onClick={() => {
                    setParentOrgNodeId(null);
                    setIsAddOrgNodeDialogOpen(true);
                  }}
                  className="rounded-xl font-bold gap-2 text-xs"
                >
                  <Plus className="h-4 w-4" /> TAMBAH JABATAN PUNCAK
                </Button>
              </div>

              <div className="overflow-x-auto pb-12 custom-scrollbar">
                <div className="min-w-[1000px] flex flex-col items-center">
                  {orgNodes.filter(n => !n.parentId).length === 0 ? (
                    <div className="text-center py-20 bg-muted/30 rounded-3xl border-2 border-dashed border-muted w-full max-w-lg">
                      <Network className="h-12 w-12 text-muted-foreground/20 mx-auto mb-4" />
                      <h3 className="text-lg font-bold text-muted-foreground">Belum Ada Struktur</h3>
                      <p className="text-sm text-muted-foreground/60 mt-1">Silakan tambahkan jabatan struktural pertama.</p>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center gap-12">
                      {orgNodes.filter(n => !n.parentId).map(rootNode => (
                        <div key={rootNode.id} className="flex flex-col items-center w-full">
                          <OrgNodeCard 
                            node={rootNode} 
                            allNodes={orgNodes} 
                            onAddSub={(id) => {
                              setParentOrgNodeId(id);
                              setIsAddOrgNodeDialogOpen(true);
                            }}
                            onEdit={(node) => {
                              setEditingOrgNode(node);
                              setIsEditOrgNodeDialogOpen(true);
                            }}
                            onDelete={handleDeleteOrgNode}
                          />
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </motion.div>
          </div>
        ) : activeTab === 'tim-kerja' ? (
          <div className="space-y-6">
            {!selectedUnitForTeams ? (
              <div className="grid gap-6 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
                {UNIT_KERJA_LIST.map((unit, index) => {
                  const unitData = groupedTeams[unit];
                  const teamCount = unitData ? (unitData.topLevel.length + Object.values(unitData.subTeams).flat().length) : 0;
                  
                  return (
                    <motion.div
                      key={unit}
                      initial={{ opacity: 0, y: 20 }}
                      animate={{ opacity: 1, y: 0 }}
                      transition={{ delay: index * 0.05 }}
                      whileHover={{ y: -8, scale: 1.02 }}
                      whileTap={{ scale: 0.98 }}
                      onClick={() => setSelectedUnitForTeams(unit)}
                      className="cursor-pointer"
                    >
                      <Card className="h-full border-none shadow-sm hover:shadow-xl transition-all duration-300 group overflow-hidden bg-white/50 backdrop-blur-sm ring-1 ring-black/5 hover:ring-primary/20">
                        <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-primary/20 via-primary to-primary/20 opacity-0 group-hover:opacity-100 transition-opacity" />
                        <CardHeader className="pt-6 pb-2 px-6">
                          <div className="flex items-center justify-between">
                            <div className="h-12 w-12 rounded-2xl bg-primary/5 flex items-center justify-center text-primary group-hover:bg-primary group-hover:text-white transition-all duration-300 shadow-inner">
                              <Building2 className="h-6 w-6" />
                            </div>
                            <div className="text-right">
                              <div className="text-2xl font-black text-primary/10 group-hover:text-primary/20 transition-colors leading-none">
                                {String(index + 1).padStart(2, '0')}
                              </div>
                            </div>
                          </div>
                        </CardHeader>
                        <CardContent className="pt-4 pb-8 px-6 flex flex-col justify-between min-h-[160px]">
                          <div>
                            <h3 className="font-extrabold text-lg text-[#1A4A9A] line-clamp-3 leading-tight group-hover:text-primary transition-colors tracking-tight">
                              {unit}
                            </h3>
                            <div className="mt-4 flex items-center gap-2">
                              <Badge variant="secondary" className="bg-primary/10 text-primary border-none font-bold px-2 py-0.5 rounded-md text-[10px]">
                                {teamCount} TIM KERJA
                              </Badge>
                            </div>
                          </div>
                          <div className="mt-8 flex items-center text-[10px] font-black uppercase tracking-widest text-primary/40 group-hover:text-primary transition-colors">
                            Buka Dashboard Unit
                            <ArrowRight className="ml-2 h-3.5 w-3.5 group-hover:translate-x-1 transition-transform" />
                          </div>
                        </CardContent>
                      </Card>
                    </motion.div>
                  );
                })}
              </div>
            ) : !selectedTeamIdForDetails ? (
              <div className="space-y-8">
                <motion.div 
                  initial={{ opacity: 0, x: -20 }}
                  animate={{ opacity: 1, x: 0 }}
                  className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 bg-white/80 backdrop-blur-md p-6 rounded-2xl shadow-sm border border-primary/10"
                >
                  <div className="flex items-center gap-4">
                    <Button 
                      variant="outline" 
                      size="sm" 
                      onClick={() => {
                        setSelectedUnitForTeams(null);
                        setSearchQuery('');
                      }}
                      className="rounded-xl h-12 w-12 p-0 border-primary/10 hover:bg-primary hover:text-white transition-all hover:scale-105 active:scale-95"
                    >
                      <ArrowLeft className="h-5 w-5" />
                    </Button>
                    <div>
                      <div className="flex items-center gap-2 mb-1">
                        <div className="h-5 w-5 rounded bg-primary/10 flex items-center justify-center">
                          <Building2 className="h-3 w-3 text-primary" />
                        </div>
                        <span className="text-[10px] font-black uppercase tracking-widest text-primary/60">Unit Kerja</span>
                      </div>
                      <h3 className="text-xl font-black text-[#1A4A9A] tracking-tight">{selectedUnitForTeams}</h3>
                      
                      {(() => {
                        const unitInfo = unitInfoList.find(u => u.unitKerja === selectedUnitForTeams);
                        if (!unitInfo && !isEditingUnitSK) {
                          return (
                            <Button 
                              variant="ghost" 
                              size="sm" 
                              onClick={() => setIsEditingUnitSK(true)}
                              className="h-6 text-[9px] font-black text-primary/60 hover:text-primary uppercase tracking-widest p-0 mt-1"
                            >
                              + TAMBAH INFO SK UNIT
                            </Button>
                          );
                        }
                        
                        if (unitInfo && !isEditingUnitSK) {
                          return (
                            <div className="flex flex-wrap items-center gap-x-4 gap-y-1 mt-1">
                              {unitInfo.nomorSK && (
                                <div className="flex items-center gap-1.5 min-w-0">
                                  <FileText className="h-3 w-3 text-primary/40 shrink-0" />
                                  <span className="text-[10px] font-bold text-primary/80 truncate">SK: {unitInfo.nomorSK}</span>
                                </div>
                              )}
                              {unitInfo.tanggalSK && (
                                <div className="flex items-center gap-1.5 min-w-0">
                                  <Calendar className="h-3 w-3 text-primary/40 shrink-0" />
                                  <span className="text-[10px] font-bold text-primary/80">
                                    {format(new Date(unitInfo.tanggalSK), 'dd MMMM yyyy', { locale: id })}
                                  </span>
                                </div>
                              )}
                              <Button 
                                variant="ghost" 
                                size="sm" 
                                onClick={() => setIsEditingUnitSK(true)}
                                className="h-5 w-5 p-0 text-primary/40 hover:text-primary"
                              >
                                <Pencil className="h-3 w-3" />
                              </Button>
                            </div>
                          );
                        }
                        return null;
                      })()}
                    </div>
                  </div>

                  <div className="flex flex-1 max-w-md mx-auto lg:mx-0 relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
                    <Input
                      placeholder="Cari tim, bagian, atau nama pegawai..."
                      className="pl-10 h-11 bg-white/50 border-primary/10 focus:border-primary/30 rounded-xl"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                    />
                  </div>
                  
                  <div className="flex items-center gap-3 justify-between sm:justify-end">
                    <Button 
                      size="sm" 
                      className="gap-2 h-11 px-6 rounded-xl font-bold shadow-lg shadow-primary/20 hover:scale-[1.02] active:scale-[0.98] transition-all"
                      onClick={() => {
                        setJenisAddTeam('Tim Kerja');
                        setIsAddTeamDialogOpen(true);
                      }}
                    >
                      <Plus className="h-4 w-4" />
                      Tambah Tim Baru
                    </Button>
                  </div>
                </motion.div>

                {groupedTeams[selectedUnitForTeams] ? (() => {
                  const unitData = groupedTeams[selectedUnitForTeams];
                  const q = searchQuery.toLowerCase();
                  
                  const filteredTopLevel = unitData.topLevel.filter(team => {
                    const matchName = team.namaTim.toLowerCase().includes(q);
                    const matchKetua = team.ketua.nama.toLowerCase().includes(q);
                    const matchAnggota = team.anggota.some(a => a.nama.toLowerCase().includes(q));
                    
                    // Also check sub-teams for matches to ensure parent is shown
                    const subs = unitData.subTeams[team.id] || [];
                    const matchSub = subs.some(sub => 
                      sub.namaTim.toLowerCase().includes(q) || 
                      sub.ketua.nama.toLowerCase().includes(q) || 
                      sub.anggota.some(a => a.nama.toLowerCase().includes(q))
                    );

                    return matchName || matchKetua || matchAnggota || matchSub;
                  });

                  if (filteredTopLevel.length === 0) {
                    return (
                      <div className="flex flex-col items-center justify-center py-20 bg-white/50 backdrop-blur-sm rounded-2xl border border-dashed border-primary/20">
                        <div className="h-16 w-16 rounded-full bg-primary/5 flex items-center justify-center mb-4">
                          <Search className="h-8 w-8 text-primary/40" />
                        </div>
                        <h3 className="text-lg font-bold text-muted-foreground">Tidak Ada Hasil</h3>
                        <p className="text-sm text-muted-foreground/60 mt-1">Coba kata kunci tim atau nama pegawai lain.</p>
                      </div>
                    );
                  }

                  return (
                    <div className="grid gap-4">
                      {filteredTopLevel.map((team, tIdx) => {
                        const subs = (unitData.subTeams[team.id] || []).filter(sub => {
                          const matchName = sub.namaTim.toLowerCase().includes(q);
                          const matchKetua = sub.ketua.nama.toLowerCase().includes(q);
                          const matchAnggota = sub.anggota.some(a => a.nama.toLowerCase().includes(q));
                          return matchName || matchKetua || matchAnggota;
                        });

                        return (
                          <motion.div 
                            key={team.id}
                            initial={{ opacity: 0, y: 10 }}
                            animate={{ opacity: 1, y: 0 }}
                            transition={{ delay: tIdx * 0.05 }}
                            className="bg-white/60 backdrop-blur-sm rounded-2xl overflow-hidden border border-primary/5 hover:border-primary/20 shadow-sm hover:shadow-md transition-all group"
                          >
                            <div 
                              className="p-5 flex items-center justify-between cursor-pointer"
                              onClick={() => setSelectedTeamIdForDetails(team.id)}
                            >
                              <div className="flex items-center gap-4">
                                <div className={cn(
                                  "h-10 w-10 rounded-xl flex items-center justify-center transition-colors",
                                  team.jenis === 'Bagian' ? "bg-amber-100 text-amber-700" : "bg-primary/10 text-primary"
                                )}>
                                  {team.jenis === 'Bagian' ? <Building2 className="h-5 w-5" /> : <Briefcase className="h-5 w-5" />}
                                </div>
                                <div>
                                  <div className="flex items-center gap-2">
                                    <h4 className="font-bold text-[#1A4A9A] tracking-tight group-hover:text-primary transition-colors">{team.namaTim}</h4>
                                    <Badge variant="outline" className="text-[8px] font-black uppercase h-4 px-1.5 border-none bg-muted text-muted-foreground">
                                      {team.jenis}
                                    </Badge>
                                  </div>
                                  <p className="text-[10px] text-muted-foreground font-medium uppercase tracking-widest mt-0.5">
                                    {team.ketua.nama} • {team.anggota.length} Anggota
                                  </p>
                                </div>
                              </div>
                              <div className="flex items-center gap-2">
                                {subs.length > 0 && (
                                  <Badge variant="secondary" className="bg-primary/5 text-primary text-[10px] font-bold border-none px-2 py-0.5">
                                    {subs.length} Sub-Tim
                                  </Badge>
                                )}
                                <div className="h-8 w-8 rounded-full flex items-center justify-center text-muted-foreground group-hover:bg-primary group-hover:text-white transition-all">
                                  <ArrowRight className="h-4 w-4" />
                                </div>
                              </div>
                            </div>

                            {subs.length > 0 && (
                              <div className="bg-primary/5 px-5 py-3 border-t border-primary/5">
                                <div className="flex flex-wrap gap-2">
                                  {subs.map(sub => (
                                    <button
                                      key={sub.id}
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        setSelectedTeamIdForDetails(sub.id);
                                      }}
                                      className="flex items-center gap-1.5 px-3 py-1.5 bg-white rounded-lg border border-primary/10 hover:border-primary/30 hover:bg-white/80 transition-all text-[10px] font-bold text-[#1A4A9A]"
                                    >
                                      <ArrowRight className="h-3 w-3 text-primary" />
                                      {sub.namaTim}
                                    </button>
                                  ))}
                                </div>
                              </div>
                            )}
                          </motion.div>
                        );
                      })}
                    </div>
                  );
                })() : null}
              </div>
            ) : (
              // Level 3: Team Details
              <div className="space-y-8">
                {(() => {
                  const team = workTeams.find(t => t.id === selectedTeamIdForDetails);
                  if (!team) return null;

                  return (
                    <>
                      <motion.div 
                        initial={{ opacity: 0, y: -10 }}
                        animate={{ opacity: 1, y: 0 }}
                        className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-[#1A4A9A] text-white p-8 rounded-3xl shadow-xl shadow-primary/20 relative overflow-hidden"
                      >
                        <div className="absolute top-0 right-0 w-64 h-64 bg-white/5 rounded-full -mr-32 -mt-32 blur-3xl" />
                        <div className="absolute bottom-0 left-0 w-48 h-48 bg-primary/20 rounded-full -ml-24 -mb-24 blur-2xl" />
                        
                        <div className="relative flex items-center gap-6">
                          <Button 
                            variant="ghost" 
                            size="sm" 
                            onClick={() => setSelectedTeamIdForDetails(null)}
                            className="rounded-2xl h-14 w-14 p-0 bg-white/10 hover:bg-white hover:text-primary transition-all text-white border border-white/20"
                          >
                            <ArrowLeft className="h-6 w-6" />
                          </Button>
                          <div>
                            <div className="flex items-center gap-2 mb-2">
                              <Badge variant="outline" className="bg-white/20 text-white border-none font-black text-[9px] px-2 py-0.5 tracking-[0.2em]">
                                {team.jenis?.toUpperCase() || 'TIM KERJA'}
                              </Badge>
                              <span className="text-[10px] text-white/60 font-bold tracking-widest uppercase italic">ID: {team.id.slice(0, 8)}</span>
                            </div>
                            <h2 className="text-3xl font-black tracking-tight">{team.namaTim}</h2>
                            <p className="text-white/70 text-sm mt-1 font-medium italic opacity-80">{team.unitKerja}</p>
                          </div>
                        </div>

                        <div className="relative flex items-center gap-1.5 self-end sm:self-center">
                          <Button 
                            variant="secondary" 
                            size="sm" 
                            className="rounded-xl font-bold bg-white text-primary hover:bg-primary-foreground h-11 px-6 shadow-lg transition-transform active:scale-95"
                            onClick={() => {
                              setEditingTeam(team);
                              setJenisEditTeam(team.jenis || 'Tim Kerja');
                              setIsEditTeamDialogOpen(true);
                            }}
                          >
                            <UserCog className="h-4 w-4 mr-2" />
                            Edit Tim
                          </Button>
                          <Button 
                            variant="secondary" 
                            size="icon" 
                            className="rounded-xl bg-white/10 text-white hover:bg-destructive hover:text-white border border-white/20 h-11 w-11 transition-all"
                            onClick={() => {
                              setDeletingInfo({ id: team.id, type: 'TIM' });
                              setIsDeleteConfirmOpen(true);
                            }}
                          >
                            <X className="h-4 w-4" />
                          </Button>
                        </div>
                      </motion.div>

                      <div className="grid gap-8 lg:grid-cols-3">
                        <div className="lg:col-span-1 space-y-6">
                          <div className="bg-white rounded-3xl shadow-sm border border-primary/5 p-6 space-y-6 relative group overflow-hidden">
                            <div className="absolute top-0 right-0 p-8 opacity-5 group-hover:scale-110 transition-transform duration-700">
                              <UserCheck className="h-32 w-32 text-primary" />
                            </div>
                            
                            <div className="flex items-center gap-2">
                              <div className="h-1.5 w-8 bg-primary rounded-full" />
                              <span className="text-[10px] font-black text-primary/60 uppercase tracking-[0.2em]">{team.jenis === 'Bagian' ? 'Pimpinan Bagian' : 'Ketua Tim'}</span>
                            </div>

                            <div className="flex flex-col items-center text-center space-y-4 relative">
                              <div className="h-28 w-28 rounded-[2rem] bg-gradient-to-br from-primary/10 to-primary/5 border-4 border-white shadow-xl flex items-center justify-center text-5xl font-black text-primary">
                                {team.ketua.nama.charAt(0)}
                              </div>
                              <div>
                                <h3 className="text-xl font-black text-[#1A4A9A] tracking-tight whitespace-normal">{team.ketua.nama}</h3>
                                <p className="text-xs text-muted-foreground font-bold mt-1 mx-auto italic whitespace-normal">{team.ketua.jabatan}</p>
                              </div>
                              <Badge className="bg-primary hover:bg-primary font-black rounded-xl px-4 py-1 tracking-widest text-[10px]">
                                {team.ketua.status.toUpperCase()}
                              </Badge>
                            </div>

                            <div className="pt-6 border-t border-dashed space-y-4">
                              <div className="flex items-center justify-center gap-8">
                                <div className="text-center">
                                  <div className="text-lg font-black text-primary">{team.tahun}</div>
                                  <div className="text-[9px] font-bold text-muted-foreground uppercase tracking-widest">Tahun Aktif</div>
                                </div>
                                <div className="h-8 w-px bg-muted" />
                                <div className="text-center">
                                  <div className="text-lg font-black text-primary">{team.status}</div>
                                  <div className="text-[9px] font-bold text-muted-foreground uppercase tracking-widest">Status Tim</div>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>

                        <div className="lg:col-span-2 space-y-6">
                          <div className="bg-white rounded-3xl shadow-sm border border-primary/5 p-6">
                            <div className="flex items-center justify-between mb-6">
                              <div className="flex items-center gap-3">
                                <div className="h-10 w-10 rounded-xl bg-primary/5 flex items-center justify-center text-primary">
                                  <Users className="h-5 w-5" />
                                </div>
                                <div>
                                  <h3 className="text-lg font-black text-[#1A4A9A] tracking-tight">Anggota Tim</h3>
                                  <p className="text-xs text-muted-foreground uppercase font-black tracking-widest opacity-60">Terdaftar {team.anggota.length} Personel</p>
                                </div>
                              </div>
                              <Button 
                                className="rounded-xl h-10 px-5 font-black text-[10px] tracking-widest shadow-lg shadow-primary/20 transition-all hover:scale-105 active:scale-95"
                                onClick={() => {
                                  setSelectedTeamForMember(team.id);
                                  setIsAddMemberDialogOpen(true);
                                }}
                              >
                                <Plus className="h-4 w-4 mr-2" />
                                TAMBAH ANGGOTA
                              </Button>
                            </div>

                            <div className="grid gap-3 sm:grid-cols-2">
                              {team.anggota.length > 0 ? (
                                team.anggota.map((member, idx) => (
                                  <motion.div 
                                    key={member.nama + idx}
                                    initial={{ opacity: 0, x: 20 }}
                                    animate={{ opacity: 1, x: 0 }}
                                    transition={{ delay: idx * 0.05 }}
                                    className="p-4 rounded-2xl bg-muted/20 border border-transparent hover:border-primary/20 hover:bg-white hover:shadow-md transition-all group/member flex items-center justify-between"
                                  >
                                    <div className="flex items-center gap-3 min-w-0">
                                      <div className="h-10 w-10 rounded-xl bg-white shadow-sm flex items-center justify-center text-xs font-black text-primary/40 group-hover/member:bg-primary group-hover/member:text-white transition-all">
                                        {idx + 1}
                                      </div>
                                      <div className="flex flex-col min-w-0">
                                        <div className="flex items-center gap-1.5 flex-wrap">
                                          <span className="text-sm font-black text-[#1A4A9A] tracking-tight whitespace-normal">{member.nama}</span>
                                          <Badge className="text-[8px] h-3.5 px-1 bg-muted-foreground/10 text-muted-foreground border-none font-bold shrink-0">
                                            {member.status}
                                          </Badge>
                                        </div>
                                        <span className="text-[10px] text-muted-foreground font-medium italic whitespace-normal mt-0.5">{member.jabatan}</span>
                                      </div>
                                    </div>
                                    
                                    <div className="flex items-center gap-1 opacity-0 group-hover/member:opacity-100 transition-all transform translate-x-2 group-hover/member:translate-x-0">
                                      <Button 
                                        variant="ghost" 
                                        size="icon" 
                                        className="h-8 w-8 rounded-lg text-muted-foreground hover:text-primary hover:bg-primary/5"
                                        onClick={() => {
                                          setEditingMemberInfo({ 
                                            teamId: team.id, 
                                            member: member, 
                                            index: idx 
                                          });
                                          setIsEditMemberDialogOpen(true);
                                        }}
                                      >
                                        <Pencil className="h-3.5 w-3.5" />
                                      </Button>
                                      <Button 
                                        variant="ghost" 
                                        size="icon" 
                                        className="h-8 w-8 rounded-lg text-muted-foreground hover:text-destructive hover:bg-destructive/5"
                                        onClick={() => {
                                          handleRemoveMember(team.id, member.nama);
                                        }}
                                      >
                                        <X className="h-3.5 w-3.5" />
                                      </Button>
                                    </div>
                                  </motion.div>
                                ))
                              ) : (
                                <div className="sm:col-span-2 flex flex-col items-center justify-center py-16 bg-muted/10 border border-dashed border-primary/10 rounded-3xl mt-4">
                                  <div className="h-12 w-12 rounded-full bg-muted flex items-center justify-center mb-4">
                                    <Users className="h-6 w-6 text-muted-foreground/30" />
                                  </div>
                                  <h4 className="text-sm font-black text-muted-foreground uppercase tracking-widest">Daftar Anggota Kosong</h4>
                                  <p className="text-[10px] font-medium text-muted-foreground/60 mt-1">Gunakan tombol di atas untuk menambahkan anggota baru.</p>
                                </div>
                              )}
                            </div>
                          </div>
                        </div>
                      </div>
                    </>
                  );
                })()}
              </div>
            )}
          </div>
        ) : activeTab === 'skp' ? (
          <div className="space-y-8">
            {/* Stats Header */}
            <div className="grid gap-6 md:grid-cols-3">
              <motion.div
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                transition={{ delay: 0.1 }}
                className="bg-white rounded-[2rem] p-6 shadow-xl shadow-primary/5 border border-primary/5 relative overflow-hidden group"
              >
                <div className="absolute top-0 right-0 p-8 opacity-10 group-hover:scale-110 transition-transform duration-500">
                  <Archive className="h-24 w-24 text-primary" />
                </div>
                <div className="relative z-10">
                  <div className="h-12 w-12 rounded-2xl bg-primary/10 flex items-center justify-center mb-4">
                    <Archive className="h-6 w-6 text-primary" />
                  </div>
                  <h3 className="text-sm font-black text-muted-foreground uppercase tracking-widest leading-none mb-2">Total Berkas</h3>
                  <div className="flex items-baseline gap-2">
                    <span className="text-4xl font-black text-[#1A4A9A] tracking-tighter">{stats.total}</span>
                    <span className="text-xs font-bold text-muted-foreground uppercase tracking-wider">Arsip</span>
                  </div>
                </div>
              </motion.div>

              <motion.div
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                transition={{ delay: 0.2 }}
                className="bg-white rounded-[2rem] p-6 shadow-xl shadow-emerald-500/5 border border-emerald-500/10 relative overflow-hidden group"
              >
                <div className="absolute top-0 right-0 p-8 opacity-10 group-hover:scale-110 transition-transform duration-500">
                  <CheckCircle2 className="h-24 w-24 text-emerald-500" />
                </div>
                <div className="relative z-10">
                  <div className="h-12 w-12 rounded-2xl bg-emerald-500/10 flex items-center justify-center mb-4">
                    <FileCheck className="h-6 w-6 text-emerald-500" />
                  </div>
                  <h3 className="text-sm font-black text-muted-foreground uppercase tracking-widest leading-none mb-2">Sudah Diambil</h3>
                  <div className="flex items-baseline gap-2">
                    <span className="text-4xl font-black text-emerald-600 tracking-tighter">{stats.sudahDiambil}</span>
                    <Badge className="bg-emerald-500/10 text-emerald-600 border-none font-black text-[10px] tracking-widest">
                      {Math.round((stats.sudahDiambil / (stats.total || 1)) * 100)}%
                    </Badge>
                  </div>
                </div>
              </motion.div>

              <motion.div
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                transition={{ delay: 0.3 }}
                className="bg-white rounded-[2rem] p-6 shadow-xl shadow-amber-500/5 border border-amber-500/10 relative overflow-hidden group"
              >
                <div className="absolute top-0 right-0 p-8 opacity-10 group-hover:scale-110 transition-transform duration-500">
                  <Clock className="h-24 w-24 text-amber-500" />
                </div>
                <div className="relative z-10">
                  <div className="h-12 w-12 rounded-2xl bg-amber-500/10 flex items-center justify-center mb-4">
                    <FileClock className="h-6 w-6 text-amber-500" />
                  </div>
                  <h3 className="text-sm font-black text-muted-foreground uppercase tracking-widest leading-none mb-2">Belum Diambil</h3>
                  <div className="flex items-baseline gap-2">
                    <span className="text-4xl font-black text-amber-600 tracking-tighter">{stats.belumDiambil}</span>
                    <span className="text-xs font-bold text-muted-foreground uppercase tracking-wider">Menunggu</span>
                  </div>
                </div>
              </motion.div>
            </div>

            {/* Main Content Area */}
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-white/80 backdrop-blur-md rounded-[2.5rem] shadow-xl shadow-primary/5 border border-primary/5 overflow-hidden"
            >
              <div className="p-8 border-b border-primary/5 bg-gradient-to-r from-primary/[0.02] to-transparent">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                  <div>
                    <h2 className="text-2xl font-black text-[#1A4A9A] tracking-tight">Arsip Berkas SKP</h2>
                    <p className="text-muted-foreground font-medium mt-1 uppercase tracking-widest text-[10px]">Manajemen Pengambilan Berkas Hasil Penilaian Kinerja</p>
                  </div>
                  <div className="flex flex-wrap items-center gap-3">
                    <div className="relative group w-full md:w-72">
                      <Search className="absolute left-4 top-1/2 -translate-y-1/2 h-4 w-4 text-primary/40 group-focus-within:text-primary transition-colors" />
                      <Input 
                        placeholder="Nama, NIP, atau Unit..." 
                        className="pl-11 h-12 rounded-2xl bg-primary/[0.03] border-none focus-visible:ring-2 focus-visible:ring-primary/20 transition-all font-medium text-sm"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                      />
                    </div>
                    <Select value={statusFilter} onValueChange={setStatusFilter}>
                      <SelectTrigger className="w-[180px] h-12 rounded-2xl bg-primary/[0.03] border-none font-bold text-xs uppercase tracking-widest px-4">
                        <div className="flex items-center gap-2">
                          <Filter className="h-4 w-4 text-primary/40" />
                          <SelectValue placeholder="Semua Status" />
                        </div>
                      </SelectTrigger>
                      <SelectContent className="rounded-2xl border-primary/5 shadow-xl">
                        <SelectItem value="all" className="font-bold text-xs uppercase tracking-widest py-3">Semua Status</SelectItem>
                        <SelectItem value="Belum Diambil" className="font-bold text-xs uppercase tracking-widest py-3 text-amber-600">Belum Diambil</SelectItem>
                        <SelectItem value="Sudah Diambil" className="font-bold text-xs uppercase tracking-widest py-3 text-emerald-600">Sudah Diambil</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                </div>
              </div>

              <div className="p-4 overflow-x-auto custom-scrollbar">
                <Table>
                  <TableHeader>
                    <TableRow className="border-none hover:bg-transparent">
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Pegawai / Identitas</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Periode</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6 text-center">Status</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Keterangan</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6 text-right">Aksi</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {filteredRecords.map((record, rIdx) => (
                      <motion.tr
                        key={record.id}
                        initial={{ opacity: 0, x: -10 }}
                        animate={{ opacity: 1, x: 0 }}
                        transition={{ delay: rIdx * 0.05 }}
                        className="group border-none hover:bg-primary/[0.02] transition-colors rounded-2xl"
                      >
                        <TableCell className="py-5 pl-4">
                          <div className="flex items-center gap-4">
                            <div className="h-12 w-12 rounded-2xl bg-gradient-to-br from-primary/5 to-primary/10 flex items-center justify-center text-[#1A4A9A] font-black group-hover:scale-110 transition-transform">
                              {record.namaPegawai.charAt(0)}
                            </div>
                            <div className="flex flex-col min-w-0">
                              <button 
                                onClick={() => {
                                  setDetailRecord(record);
                                  setIsDetailDialogOpen(true);
                                }}
                                className="text-left group/name"
                              >
                                <span className="font-black text-sm text-[#1A4A9A] group-hover/name:text-primary transition-colors block leading-tight">{record.namaPegawai}</span>
                              </button>
                              <div className="flex items-center gap-2 mt-1">
                                <span className="text-[10px] font-bold text-muted-foreground uppercase tracking-widest break-all line-clamp-1">{record.jabatan}</span>
                              </div>
                              <span className="text-[9px] font-mono text-primary/40 font-bold mt-1">NIP. {record.nip}</span>
                            </div>
                          </div>
                        </TableCell>
                        <TableCell>
                          <div className="flex flex-col">
                            <span className="text-sm font-black text-[#1A4A9A] tracking-tighter">{record.tahun}</span>
                            <Badge className="w-fit text-[8px] h-4 px-1.5 font-black uppercase tracking-widest bg-primary/10 text-primary border-none">
                              {record.periode}
                            </Badge>
                          </div>
                        </TableCell>
                        <TableCell className="text-center">
                          <div className="flex flex-col items-center gap-1.5">
                            <Badge className={cn(
                              "rounded-full px-4 py-1 text-[9px] font-black uppercase tracking-widest border-none ring-1",
                              record.status === 'Sudah Diambil' 
                                ? "bg-emerald-50 text-emerald-600 ring-emerald-500/20" 
                                : "bg-amber-50 text-amber-600 ring-amber-500/20"
                            )}>
                              {record.status}
                            </Badge>
                          </div>
                        </TableCell>
                        <TableCell>
                          {record.status === 'Sudah Diambil' ? (
                            <div className="bg-primary/[0.03] p-3 rounded-2xl border border-primary/5">
                              <div className="flex items-center gap-2 mb-1.5">
                                <UserCheck className="h-3 w-3 text-primary/40" />
                                <span className="text-[10px] font-black text-primary/80 uppercase">{record.pengambil}</span>
                              </div>
                              <div className="flex items-center gap-2">
                                <Calendar className="h-3 w-3 text-primary/40" />
                                <span className="text-[9px] font-bold text-primary/60">
                                  {record.tanggalAmbil && format(new Date(record.tanggalAmbil), 'dd MMM yyyy, HH:mm', { locale: id })}
                                </span>
                              </div>
                            </div>
                          ) : (
                            <div className="flex items-center gap-2 text-muted-foreground/40 italic text-[10px] font-medium justify-center">
                              <Clock className="h-3.5 w-3.5" />
                              Belum ada data pengambilan
                            </div>
                          )}
                        </TableCell>
                        <TableCell className="text-right pr-4">
                          <div className="flex items-center justify-end gap-2">
                            {record.status === 'Belum Diambil' ? (
                              <Button 
                                size="sm" 
                                className="h-9 rounded-xl bg-emerald-500 hover:bg-emerald-600 text-white font-black text-[10px] tracking-widest px-4 gap-2 shadow-lg shadow-emerald-500/20"
                                onClick={() => {
                                  setSelectedRecord(record);
                                  setIsTakeDialogOpen(true);
                                }}
                              >
                                <FileCheck className="h-3.5 w-3.5" /> AMBIL
                              </Button>
                            ) : (
                              <Button 
                                size="icon" 
                                variant="ghost" 
                                className="h-9 w-9 rounded-xl hover:bg-primary/5 text-primary/40 hover:text-primary transition-all opacity-0 group-hover:opacity-100"
                                onClick={() => {
                                  setEditingRecord(record);
                                  setIsEditDialogOpen(true);
                                }}
                              >
                                <Pencil className="h-4 w-4" />
                              </Button>
                            )}
                            <Button 
                              size="icon" 
                              variant="ghost" 
                              className="h-9 w-9 rounded-xl hover:bg-destructive/5 text-destructive/40 hover:text-destructive transition-all opacity-0 group-hover:opacity-100"
                              onClick={() => handleDeleteRecord(record.id)}
                            >
                              <Trash2 className="h-4 w-4" />
                            </Button>
                          </div>
                        </TableCell>
                      </motion.tr>
                    ))}
                  </TableBody>
                </Table>
              </div>
            </motion.div>
          </div>
        ) : activeTab === 'skp-monitoring' ? (
          <div className="space-y-8">
            {/* Monitoring Stats - Bento Grid Style */}
            {!monitoringSelectedEmployee && (
              <div className="grid gap-6 md:grid-cols-4">
                <motion.div initial={{ opacity: 0, scale: 0.9 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.1 }}>
                  <div className="bg-emerald-50/50 backdrop-blur-sm border border-emerald-100 p-6 rounded-[2rem] hover:shadow-lg hover:shadow-emerald-500/5 transition-all group">
                    <div className="flex items-center justify-between mb-4">
                      <div className="h-10 w-10 rounded-xl bg-emerald-100 flex items-center justify-center text-emerald-600">
                        <TrendingUp className="h-5 w-5" />
                      </div>
                      <span className="text-[10px] font-black text-emerald-600/60 uppercase tracking-[0.2em]">Rating A</span>
                    </div>
                    <div className="text-4xl font-black text-emerald-900 tracking-tighter">{skpMonitoringStats.sangatBaik}</div>
                    <p className="text-xs font-bold text-emerald-700 uppercase tracking-widest mt-1">Sangat Baik</p>
                  </div>
                </motion.div>

                <motion.div initial={{ opacity: 0, scale: 0.9 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.2 }}>
                  <div className="bg-blue-50/50 backdrop-blur-sm border border-blue-100 p-6 rounded-[2rem] hover:shadow-lg hover:shadow-blue-500/5 transition-all group">
                    <div className="flex items-center justify-between mb-4">
                      <div className="h-10 w-10 rounded-xl bg-blue-100 flex items-center justify-center text-blue-600">
                        <CheckCircle2 className="h-5 w-5" />
                      </div>
                      <span className="text-[10px] font-black text-blue-600/60 uppercase tracking-[0.2em]">Rating B</span>
                    </div>
                    <div className="text-4xl font-black text-blue-900 tracking-tighter">{skpMonitoringStats.baik}</div>
                    <p className="text-xs font-bold text-blue-700 uppercase tracking-widest mt-1">Baik</p>
                  </div>
                </motion.div>

                <motion.div initial={{ opacity: 0, scale: 0.9 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.3 }}>
                  <div className="bg-amber-50/50 backdrop-blur-sm border border-amber-100 p-6 rounded-[2rem] hover:shadow-lg hover:shadow-amber-500/5 transition-all group">
                    <div className="flex items-center justify-between mb-4">
                      <div className="h-10 w-10 rounded-xl bg-amber-100 flex items-center justify-center text-amber-600">
                        <AlertTriangle className="h-5 w-5" />
                      </div>
                      <span className="text-[10px] font-black text-amber-600/60 uppercase tracking-[0.2em]">Rating C</span>
                    </div>
                    <div className="text-4xl font-black text-amber-900 tracking-tighter">{skpMonitoringStats.butuhPerbaikan}</div>
                    <p className="text-xs font-bold text-amber-700 uppercase tracking-widest mt-1">Butuh Perbaikan</p>
                  </div>
                </motion.div>

                <motion.div initial={{ opacity: 0, scale: 0.9 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.4 }}>
                  <div className="bg-red-50/50 backdrop-blur-sm border border-red-100 p-6 rounded-[2rem] hover:shadow-lg hover:shadow-red-500/5 transition-all group">
                    <div className="flex items-center justify-between mb-4">
                      <div className="h-10 w-10 rounded-xl bg-red-100 flex items-center justify-center text-red-600">
                        <XCircle className="h-5 w-5" />
                      </div>
                      <span className="text-[10px] font-black text-red-600/60 uppercase tracking-[0.2em]">Rating D/E</span>
                    </div>
                    <div className="text-4xl font-black text-red-900 tracking-tighter">{skpMonitoringStats.kurang}</div>
                    <p className="text-xs font-bold text-red-700 uppercase tracking-widest mt-1">Kurang</p>
                  </div>
                </motion.div>
              </div>
            )}

            {/* Monitoring List / Profile View */}
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-white rounded-[2.5rem] shadow-xl shadow-primary/5 border border-primary/5 overflow-hidden"
            >
              <div className="p-8 border-b border-primary/5 bg-gradient-to-r from-primary/[0.01] to-transparent">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                  <div className="flex items-center gap-6">
                    {monitoringSelectedEmployee && (
                      <Button 
                        variant="ghost" 
                        size="icon" 
                        className="rounded-2xl h-12 w-12 bg-primary/5 hover:bg-primary/10 text-primary transition-all"
                        onClick={() => setMonitoringSelectedEmployee(null)}
                      >
                        <ArrowLeft className="h-5 w-5" />
                      </Button>
                    )}
                    <div>
                      <h2 className="text-2xl font-black text-[#1A4A9A] tracking-tight">
                        {monitoringSelectedEmployee 
                          ? `Riwayat SKP: ${selectedEmployeeHistory[0]?.namaPegawai || 'Pegawai'}` 
                          : 'Monitoring Rating SKP'}
                      </h2>
                      <p className="text-muted-foreground font-medium mt-1 uppercase tracking-widest text-[10px]">
                        {monitoringSelectedEmployee 
                          ? `NIP. ${selectedEmployeeHistory[0]?.nip || '-'} • ${selectedEmployeeHistory[0]?.unitKerja || '-'}`
                          : 'Visualisasi Capaian dan Grafik Penilaian Kinerja Pegawai'}
                      </p>
                    </div>
                  </div>
                  {!monitoringSelectedEmployee && (
                    <div className="relative group w-full md:w-80">
                      <Search className="absolute left-4 top-1/2 -translate-y-1/2 h-4 w-4 text-primary/40 group-focus-within:text-primary transition-colors" />
                      <Input 
                        placeholder="Cari Pegawai..." 
                        className="pl-11 h-12 rounded-2xl bg-primary/[0.03] border-none focus-visible:ring-2 focus-visible:ring-primary/20 transition-all font-bold text-xs uppercase tracking-widest"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                      />
                    </div>
                  )}
                </div>
              </div>

              <div className="p-4 overflow-x-auto custom-scrollbar">
                {!monitoringSelectedEmployee ? (
                  <Table>
                    <TableHeader>
                      <TableRow className="border-none hover:bg-transparent">
                        <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Pegawai</TableHead>
                        <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Unit Kerja</TableHead>
                        <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Predikat Terakhir</TableHead>
                        <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6 text-right">Aksi</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {uniqueEmployeesForMonitoring.map((emp, eIdx) => (
                        <motion.tr 
                          key={emp.nip || emp.namaPegawai} 
                          initial={{ opacity: 0, y: 10 }}
                          animate={{ opacity: 1, y: 0 }}
                          transition={{ delay: eIdx * 0.05 }}
                          className="group border-none hover:bg-primary/[0.02] transition-colors cursor-pointer" 
                          onClick={() => setMonitoringSelectedEmployee(emp.nip || emp.namaPegawai)}
                        >
                          <TableCell className="py-5">
                            <div className="flex items-center gap-4">
                              <div className="h-10 w-10 rounded-xl bg-gradient-to-br from-primary/5 to-primary/10 flex items-center justify-center text-[#1A4A9A] font-black uppercase text-sm">
                                {emp.namaPegawai.charAt(0)}
                              </div>
                              <div className="flex flex-col">
                                <span className="font-black text-sm text-[#1A4A9A] leading-tight group-hover:text-primary transition-colors">{emp.namaPegawai}</span>
                                <span className="text-[10px] font-mono font-bold text-primary/40 mt-1">NIP. {emp.nip}</span>
                              </div>
                            </div>
                          </TableCell>
                          <TableCell>
                            <span className="text-[10px] font-black uppercase tracking-widest text-muted-foreground/60">{emp.unitKerja}</span>
                          </TableCell>
                          <TableCell>
                            {emp.latestPredikat ? (
                              <div className="flex items-center gap-2">
                                <div className={cn(
                                  "h-2 w-2 rounded-full",
                                  emp.latestPredikat === 'Sangat Baik' ? "bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.4)]" :
                                  emp.latestPredikat === 'Baik' ? "bg-blue-500 shadow-[0_0_8px_rgba(59,130,246,0.4)]" :
                                  emp.latestPredikat === 'Butuh Perbaikan' ? "bg-amber-500 shadow-[0_0_8px_rgba(245,158,11,0.4)]" : 
                                  "bg-red-500 shadow-[0_0_8px_rgba(239,68,68,0.4)]"
                                )} />
                                <span className="font-black text-xs uppercase tracking-tight text-[#1A4A9A]">{emp.latestPredikat}</span>
                              </div>
                            ) : (
                              <span className="text-[10px] text-muted-foreground font-medium italic opacity-40">Belum ada data</span>
                            )}
                          </TableCell>
                          <TableCell className="text-right">
                            <Button size="sm" variant="ghost" className="rounded-xl h-9 font-black text-[9px] uppercase tracking-[0.2em] text-primary/60 hover:text-primary hover:bg-primary/5 px-4">
                              Detail Riwayat
                            </Button>
                          </TableCell>
                        </motion.tr>
                      ))}
                    </TableBody>
                  </Table>
                ) : (
                  <Table>
                    <TableHeader>
                      <TableRow className="border-none hover:bg-transparent">
                        <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6 text-center">Tahun</TableHead>
                        <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Hasil Kerja</TableHead>
                        <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Perilaku</TableHead>
                        <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Predikat Akhir</TableHead>
                        <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6 text-right">Aksi</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {selectedEmployeeHistory.map((record, rIdx) => (
                        <motion.tr 
                          key={record.id} 
                          initial={{ opacity: 0, x: -10 }}
                          animate={{ opacity: 1, x: 0 }}
                          transition={{ delay: rIdx * 0.1 }}
                          className="group border-none hover:bg-primary/[0.02] transition-colors"
                        >
                          <TableCell className="text-center py-6">
                            <div className="flex flex-col items-center">
                              <span className="text-lg font-black text-[#1A4A9A] tracking-tighter leading-none">{record.tahun}</span>
                              <Badge className="mt-1 h-4 px-1.5 text-[8px] font-black uppercase tracking-widest bg-primary/5 text-primary border-none">
                                {record.periode}
                              </Badge>
                            </div>
                          </TableCell>
                          <TableCell>
                            {record.ratingHasilKerja ? (
                              <Badge className="bg-emerald-50 text-emerald-600 border-none text-[9px] font-black uppercase tracking-widest py-1 px-3">
                                {record.ratingHasilKerja}
                              </Badge>
                            ) : (
                              <span className="text-[10px] text-muted-foreground opacity-40 italic">Kosong</span>
                            )}
                          </TableCell>
                          <TableCell>
                            {record.ratingHasilPerilaku ? (
                              <Badge className="bg-indigo-50 text-indigo-600 border-none text-[9px] font-black uppercase tracking-widest py-1 px-3">
                                {record.ratingHasilPerilaku}
                              </Badge>
                            ) : (
                              <span className="text-[10px] text-muted-foreground opacity-40 italic">Kosong</span>
                            )}
                          </TableCell>
                          <TableCell>
                            {record.predikatKinerja ? (
                              <div className="flex items-center gap-2">
                                <div className={cn(
                                  "h-2 w-2 rounded-full",
                                  record.predikatKinerja === 'Sangat Baik' ? "bg-emerald-500" :
                                  record.predikatKinerja === 'Baik' ? "bg-blue-500" :
                                  record.predikatKinerja === 'Butuh Perbaikan' ? "bg-amber-500" : "bg-red-500"
                                )} />
                                <span className="font-black text-xs uppercase tracking-tight text-[#1A4A9A]">{record.predikatKinerja}</span>
                              </div>
                            ) : (
                              <span className="text-[10px] text-muted-foreground opacity-40 font-bold">Menunggu...</span>
                            )}
                          </TableCell>
                          <TableCell className="text-right">
                            <div className="flex justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                              <Button 
                                size="icon" 
                                variant="ghost" 
                                className="h-9 w-9 rounded-xl hover:bg-primary/5 text-primary/40 hover:text-primary"
                                onClick={() => {
                                  setEditingMonitoringRecord(record);
                                  setIsEditMonitoringDialogOpen(true);
                                }}
                              >
                                <Pencil className="h-4 w-4" />
                              </Button>
                              <Button 
                                size="icon" 
                                variant="ghost" 
                                className="h-9 w-9 rounded-xl hover:bg-destructive/5 text-destructive/40 hover:text-destructive"
                                onClick={() => {
                                  setDeletingInfo({ id: record.id!, type: 'SKP_MONITORING' });
                                  setIsDeleteConfirmOpen(true);
                                }}
                              >
                                <Trash2 className="h-4 w-4" />
                              </Button>
                            </div>
                          </TableCell>
                        </motion.tr>
                      ))}
                    </TableBody>
                  </Table>
                )}
              </div>
            </motion.div>
          </div>
        ) : activeTab === 'plt-plh' ? (
          <div className="space-y-8">
            {/* PLT/PLH Stats Grid - Bento Style */}
            <div className="grid gap-6 md:grid-cols-2 lg:grid-cols-4">
              <motion.div initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.1 }}>
                <div 
                  className={cn(
                    "bg-white rounded-[2rem] p-6 shadow-xl transition-all cursor-pointer border border-primary/5 relative overflow-hidden group",
                    statusFilterPlt === 'all' ? "ring-2 ring-primary ring-offset-4" : "shadow-primary/5"
                  )}
                  onClick={() => setStatusFilterPlt('all')}
                >
                  <div className="absolute top-0 right-0 p-6 opacity-5 group-hover:scale-110 transition-transform duration-500">
                    <UserCog className="h-20 w-20 text-primary" />
                  </div>
                  <div className="relative z-10">
                    <div className="h-12 w-12 rounded-2xl bg-primary/10 flex items-center justify-center mb-4 text-primary">
                      <UserCog className="h-6 w-6" />
                    </div>
                    <h3 className="text-[10px] font-black text-muted-foreground uppercase tracking-widest leading-none mb-2">Total PLT</h3>
                    <div className="flex items-baseline gap-2">
                      <span className="text-4xl font-black text-[#1A4A9A] tracking-tighter">{pltPlhStats.totalPlt}</span>
                    </div>
                  </div>
                </div>
              </motion.div>

              <motion.div initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.2 }}>
                <div 
                  className={cn(
                    "bg-white rounded-[2rem] p-6 shadow-xl transition-all cursor-pointer border border-emerald-500/5 relative overflow-hidden group",
                    statusFilterPlt === 'Aktif' ? "ring-2 ring-emerald-500 ring-offset-4" : "shadow-emerald-500/5"
                  )}
                  onClick={() => setStatusFilterPlt(statusFilterPlt === 'Aktif' ? 'all' : 'Aktif')}
                >
                  <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform duration-500">
                    <CheckCircle2 className="h-20 w-20 text-emerald-500" />
                  </div>
                  <div className="relative z-10">
                    <div className="h-12 w-12 rounded-2xl bg-emerald-500/10 flex items-center justify-center mb-4 text-emerald-600">
                      <CheckCircle2 className="h-6 w-6" />
                    </div>
                    <h3 className="text-[10px] font-black text-muted-foreground uppercase tracking-widest leading-none mb-2">PLT Aktif</h3>
                    <div className="flex items-baseline gap-2">
                      <span className="text-4xl font-black text-emerald-600 tracking-tighter">{pltPlhStats.activePlt}</span>
                    </div>
                  </div>
                </div>
              </motion.div>

              <motion.div initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.3 }}>
                <div 
                  className={cn(
                    "bg-white rounded-[2rem] p-6 shadow-xl transition-all cursor-pointer border border-blue-500/5 relative overflow-hidden group",
                    statusFilterPlh === 'all' ? "ring-2 ring-blue-500 ring-offset-4" : "shadow-blue-500/5"
                  )}
                  onClick={() => setStatusFilterPlh('all')}
                >
                  <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform duration-500">
                    <UserCheck className="h-20 w-20 text-blue-500" />
                  </div>
                  <div className="relative z-10">
                    <div className="h-12 w-12 rounded-2xl bg-blue-500/10 flex items-center justify-center mb-4 text-blue-600">
                      <UserCheck className="h-6 w-6" />
                    </div>
                    <h3 className="text-[10px] font-black text-muted-foreground uppercase tracking-widest leading-none mb-2">Total PLH</h3>
                    <div className="flex items-baseline gap-2">
                      <span className="text-4xl font-black text-blue-600 tracking-tighter">{pltPlhStats.totalPlh}</span>
                    </div>
                  </div>
                </div>
              </motion.div>

              <motion.div initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.4 }}>
                <div 
                  className={cn(
                    "bg-white rounded-[2rem] p-6 shadow-xl transition-all cursor-pointer border border-sky-500/5 relative overflow-hidden group",
                    statusFilterPlh === 'Aktif' ? "ring-2 ring-sky-500 ring-offset-4" : "shadow-sky-500/5"
                  )}
                  onClick={() => setStatusFilterPlh(statusFilterPlh === 'Aktif' ? 'all' : 'Aktif')}
                >
                  <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform duration-500">
                    <Clock className="h-20 w-20 text-sky-500" />
                  </div>
                  <div className="relative z-10">
                    <div className="h-12 w-12 rounded-2xl bg-sky-500/10 flex items-center justify-center mb-4 text-sky-600">
                      <Clock className="h-6 w-6" />
                    </div>
                    <h3 className="text-[10px] font-black text-muted-foreground uppercase tracking-widest leading-none mb-2">PLH Aktif</h3>
                    <div className="flex items-baseline gap-2">
                      <span className="text-4xl font-black text-sky-600 tracking-tighter">{pltPlhStats.activePlh}</span>
                    </div>
                  </div>
                </div>
              </motion.div>
            </div>

            {/* PLT/PLH Chart Section */}
            <motion.div 
              initial={{ opacity: 0, y: 20 }} 
              animate={{ opacity: 1, y: 0 }} 
              transition={{ delay: 0.5 }}
              className="bg-white rounded-[2.5rem] shadow-xl shadow-primary/5 border border-primary/5 overflow-hidden"
            >
              <div className="p-8 border-b border-primary/5 bg-gradient-to-r from-primary/[0.01] to-transparent">
                <h3 className="text-xl font-black text-[#1A4A9A] tracking-tight">Tren Penugasan Tahunan</h3>
                <p className="text-[10px] font-bold text-muted-foreground uppercase tracking-widest mt-1">Perbandingan Distribusi Beban Kerja PLT & PLH Berdasarkan Tahun SK</p>
              </div>
              <div className="p-8">
                <div className="h-[350px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={chartDataPltPlh}
                      margin={{ top: 20, right: 30, left: 10, bottom: 5 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E2E8F0" />
                      <XAxis 
                        dataKey="year" 
                        axisLine={false} 
                        tickLine={false} 
                        tick={{ fill: '#64748B', fontSize: 11, fontWeight: 700 }}
                        dy={10}
                      />
                      <YAxis 
                        axisLine={false} 
                        tickLine={false} 
                        tick={{ fill: '#64748B', fontSize: 11, fontWeight: 700 }}
                      />
                      <Tooltip 
                        cursor={{ fill: '#F1F5F9' }}
                        contentStyle={{ 
                          borderRadius: '1.5rem', 
                          border: '1px solid rgba(26,74,154,0.1)', 
                          boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1)',
                          backdropFilter: 'blur(8px)',
                          background: 'rgba(255,255,255,0.9)'
                        }}
                      />
                      <Legend 
                        verticalAlign="top" 
                        align="right" 
                        iconType="circle" 
                        wrapperStyle={{ paddingBottom: '30px', fontSize: '10px', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.1em' }} 
                      />
                      <Bar 
                        dataKey="PLT" 
                        fill="#1A4A9A" 
                        radius={[6, 6, 0, 0]} 
                        barSize={32}
                      />
                      <Bar 
                        dataKey="PLH" 
                        fill="#60A5FA" 
                        radius={[6, 6, 0, 0]} 
                        barSize={32}
                      />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </motion.div>

            {/* PLT Table */}
            <motion.div 
              initial={{ opacity: 0, y: 20 }} 
              animate={{ opacity: 1, y: 0 }} 
              transition={{ delay: 0.6 }}
              className="bg-white rounded-[2.5rem] shadow-xl shadow-primary/5 border border-primary/5 overflow-hidden"
            >
              <div className="p-8 border-b border-primary/5 bg-gradient-to-r from-primary/[0.01] to-transparent">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                  <div>
                    <h3 className="text-xl font-black text-[#1A4A9A] tracking-tight">Monitoring Pelaksana Tugas (PLT)</h3>
                    <p className="text-[10px] font-bold text-muted-foreground uppercase tracking-widest mt-1">Daftar Penugasan Jabatan Sementara Seluruh Unit Kerja BSKJI</p>
                  </div>
                  <div className="flex flex-wrap items-center gap-3">
                    <div className="relative group w-full md:w-64">
                      <Search className="absolute left-4 top-1/2 -translate-y-1/2 h-4 w-4 text-primary/40 group-focus-within:text-primary transition-colors" />
                      <Input 
                        placeholder="Nama atau NIP..." 
                        className="pl-11 h-12 rounded-2xl bg-primary/[0.03] border-none focus-visible:ring-2 focus-visible:ring-primary/20 transition-all font-bold text-xs uppercase tracking-widest"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                      />
                    </div>
                    <Select value={yearFilterPltPlh} onValueChange={setYearFilterPltPlh}>
                      <SelectTrigger className="w-full md:w-[130px] h-12 rounded-2xl bg-primary/[0.03] border-none font-bold text-[10px] uppercase tracking-widest px-4">
                        <div className="flex items-center gap-2">
                          <Filter className="h-4 w-4 text-primary/40" />
                          <SelectValue placeholder="Tahun" />
                        </div>
                      </SelectTrigger>
                      <SelectContent className="rounded-2xl border-primary/5 shadow-xl">
                        <SelectItem value="all" className="font-bold text-[10px] uppercase tracking-widest py-3">Semua Tahun</SelectItem>
                        {Array.from({ length: 5 }, (_, i) => new Date().getFullYear() - i).map(year => (
                          <SelectItem key={year} value={year.toString()} className="font-bold text-[10px] uppercase tracking-widest py-3">{year}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                </div>
              </div>
              <div className="p-4 overflow-x-auto custom-scrollbar">
                <Table>
                  <TableHeader>
                    <TableRow className="border-none hover:bg-transparent">
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Pegawai / Identitas</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Penugasan PLT</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Unit Asal & Tugas</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6">Status</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-primary/40 py-6 text-right">Aksi</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {filteredPltRecords.length > 0 ? (
                      filteredPltRecords.map((record, rIdx) => (
                        <motion.tr 
                          key={record.id} 
                          initial={{ opacity: 0, x: -10 }}
                          animate={{ opacity: 1, x: 0 }}
                          transition={{ delay: rIdx * 0.05 }}
                          className="group border-none hover:bg-primary/[0.02] transition-colors"
                        >
                          <TableCell className="py-5">
                            <div className="flex items-center gap-4">
                              <div className="h-10 w-10 rounded-xl bg-gradient-to-br from-primary/5 to-primary/10 flex items-center justify-center text-[#1A4A9A] font-black uppercase text-sm">
                                {record.namaPegawai.charAt(0)}
                              </div>
                              <div className="flex flex-col">
                                <span className="font-black text-sm text-[#1A4A9A] group-hover:text-primary transition-colors block leading-tight">{record.namaPegawai}</span>
                                <span className="text-[9px] font-mono text-primary/40 font-bold mt-1">NIP. {record.nip}</span>
                                <span className="text-[9px] font-bold text-muted-foreground/60 uppercase tracking-tighter mt-0.5">{record.jabatanAsli}</span>
                              </div>
                            </div>
                          </TableCell>
                          <TableCell>
                            <div className="flex flex-col">
                              <span className="font-black text-sm text-primary tracking-tight">{record.jabatanPlt}</span>
                              <div className="flex items-center gap-2 mt-1">
                                <Badge className="text-[8px] font-black uppercase tracking-widest bg-primary/5 text-primary border-none px-2 h-4">SK: {record.noSk}</Badge>
                              </div>
                            </div>
                          </TableCell>
                          <TableCell>
                            <div className="flex flex-col gap-1.5">
                              <div className="flex items-center gap-2 text-[10px] font-bold">
                                <Home className="h-3 w-3 text-emerald-500" />
                                <span className="text-emerald-600 uppercase tracking-tight">{record.unitKerja}</span>
                              </div>
                              <div className="flex items-center gap-2 text-[10px] font-bold">
                                <ArrowRight className="h-3 w-3 text-primary" />
                                <span className="text-primary uppercase tracking-tight">{record.unitKerjaTugas}</span>
                              </div>
                            </div>
                          </TableCell>
                          <TableCell>
                            <div className="flex flex-col gap-1.5">
                              <Badge className={cn(
                                "w-fit rounded-full px-3 py-1 text-[9px] font-black uppercase tracking-widest border-none",
                                record.status === 'Aktif' ? "bg-emerald-500 text-white" : "bg-muted text-muted-foreground"
                              )}>
                                {record.status}
                              </Badge>
                              <div className="flex items-center gap-1.5 text-[9px] font-bold text-muted-foreground">
                                <Calendar className="h-2.5 w-2.5" />
                                {format(new Date(record.tglMulai), 'dd/MM/yy')}
                                {record.tglSelesai && ` - ${format(new Date(record.tglSelesai), 'dd/MM/yy')}`}
                              </div>
                            </div>
                          </TableCell>
                          <TableCell className="text-right">
                            <div className="flex justify-end gap-2 opacity-0 group-hover:opacity-100 transition-all transform translate-x-2 group-hover:translate-x-0">
                              <Button 
                                size="icon" 
                                variant="ghost" 
                                className="h-9 w-9 rounded-xl hover:bg-primary/5 text-primary/40 hover:text-primary"
                                onClick={() => {
                                  setEditingPltRecord(record);
                                  setIsEditPltPlhDialogOpen(true);
                                }}
                              >
                                <Pencil className="h-4 w-4" />
                              </Button>
                              <Button 
                                size="icon" 
                                variant="ghost" 
                                className="h-9 w-9 rounded-xl hover:bg-destructive/5 text-destructive/40 hover:text-destructive"
                                onClick={() => handleDeletePltPlhRecord(record.id, 'PLT')}
                              >
                                <Trash2 className="h-4 w-4" />
                              </Button>
                            </div>
                          </TableCell>
                        </motion.tr>
                      ))
                    ) : (
                      <TableRow className="hover:bg-transparent">
                        <TableCell colSpan={5} className="h-32 text-center font-black text-[10px] uppercase tracking-[0.2em] text-primary/20">
                          Tidak ada data penugasan PLT ditemukan.
                        </TableCell>
                      </TableRow>
                    )}
                  </TableBody>
                </Table>
              </div>
            </motion.div>

            {/* PLH Table */}
            <motion.div 
              initial={{ opacity: 0, y: 20 }} 
              animate={{ opacity: 1, y: 0 }} 
              transition={{ delay: 0.7 }}
              className="bg-white rounded-[2.5rem] shadow-xl shadow-blue-500/5 border border-blue-500/10 overflow-hidden"
            >
              <div className="p-8 border-b border-blue-500/5 bg-gradient-to-r from-blue-500/[0.01] to-transparent">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                  <div>
                    <h3 className="text-xl font-black text-blue-900 tracking-tight">Monitoring Pelaksana Harian (PLH)</h3>
                    <p className="text-[10px] font-bold text-muted-foreground uppercase tracking-widest mt-1">Daftar Penugasan Jabatan Sementara Karena Pejabat Berhalangan</p>
                  </div>
                  <div className="flex flex-wrap items-center gap-3">
                    <div className="relative group w-full md:w-64">
                      <Search className="absolute left-4 top-1/2 -translate-y-1/2 h-4 w-4 text-blue-500/40 group-focus-within:text-blue-500 transition-colors" />
                      <Input 
                        placeholder="Nama atau NIP..." 
                        className="pl-11 h-12 rounded-2xl bg-blue-500/[0.03] border-none focus-visible:ring-2 focus-visible:ring-blue-500/20 transition-all font-bold text-xs uppercase tracking-widest"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                      />
                    </div>
                  </div>
                </div>
              </div>
              <div className="p-4 overflow-x-auto custom-scrollbar">
                <Table>
                  <TableHeader>
                    <TableRow className="border-none hover:bg-transparent">
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-blue-500/40 py-6">Pegawai / Identitas</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-blue-500/40 py-6">Penugasan PLH</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-blue-500/40 py-6">Jabatan Digantikan</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-blue-500/40 py-6">Periode & Status</TableHead>
                      <TableHead className="font-black text-[10px] uppercase tracking-[0.2em] text-blue-500/40 py-6 text-right">Aksi</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {filteredPlhRecords.length > 0 ? (
                      filteredPlhRecords.map((record, rIdx) => (
                        <motion.tr 
                          key={record.id} 
                          initial={{ opacity: 0, x: -10 }}
                          animate={{ opacity: 1, x: 0 }}
                          transition={{ delay: rIdx * 0.05 }}
                          className="group border-none hover:bg-blue-500/[0.02] transition-colors"
                        >
                          <TableCell className="py-5">
                            <div className="flex items-center gap-4">
                              <div className="h-10 w-10 rounded-xl bg-gradient-to-br from-blue-500/5 to-blue-500/10 flex items-center justify-center text-blue-700 font-black uppercase text-sm">
                                {record.namaPegawai.charAt(0)}
                              </div>
                              <div className="flex flex-col">
                                <span className="font-black text-sm text-blue-900 group-hover:text-blue-600 transition-colors block leading-tight">{record.namaPegawai}</span>
                                <span className="text-[9px] font-mono text-blue-500/40 font-bold mt-1">NIP. {record.nip}</span>
                              </div>
                            </div>
                          </TableCell>
                          <TableCell>
                            <div className="flex flex-col">
                              <span className="font-black text-sm text-blue-700 tracking-tight">{record.jabatanPlh}</span>
                              <Badge className="w-fit mt-1 text-[8px] font-black uppercase tracking-widest bg-blue-500/5 text-blue-600 border-none px-2 h-4">SK: {record.noSk}</Badge>
                            </div>
                          </TableCell>
                          <TableCell>
                            <div className="flex flex-col bg-blue-50/50 p-2.5 rounded-xl border border-blue-100">
                              <span className="text-[10px] font-black text-blue-900 uppercase leading-tight">{record.pegawaiDigantikan}</span>
                              <div className="flex items-center gap-1.5 mt-1.5">
                                <Building2 className="h-2.5 w-2.5 text-blue-400" />
                                <span className="text-[9px] font-bold text-blue-500/60 uppercase tracking-tighter">{record.unitKerjaTugas}</span>
                              </div>
                            </div>
                          </TableCell>
                          <TableCell>
                            <div className="flex flex-col gap-1.5">
                              <Badge className={cn(
                                "w-fit rounded-full px-3 py-1 text-[9px] font-black uppercase tracking-widest border-none",
                                record.status === 'Aktif' ? "bg-blue-500 text-white shadow-[0_0_8px_rgba(59,130,246,0.3)]" : "bg-muted text-muted-foreground"
                              )}>
                                {record.status}
                              </Badge>
                              <div className="flex items-center gap-1.5 text-[9px] font-bold text-muted-foreground">
                                <Calendar className="h-2.5 w-2.5" />
                                {format(new Date(record.tglMulai), 'dd/MM/yy')} - {record.tglSelesai ? format(new Date(record.tglSelesai), 'dd/MM/yy') : 'Selesai'}
                              </div>
                            </div>
                          </TableCell>
                          <TableCell className="text-right">
                            <div className="flex justify-end gap-2 opacity-0 group-hover:opacity-100 transition-all translate-x-2 group-hover:translate-x-0">
                              <Button 
                                size="icon" 
                                variant="ghost" 
                                className="h-9 w-9 rounded-xl hover:bg-blue-500/5 text-blue-500/40 hover:text-blue-500"
                                onClick={() => {
                                  setEditingPlhRecord(record);
                                  setIsEditPltPlhDialogOpen(true);
                                }}
                              >
                                <Pencil className="h-4 w-4" />
                              </Button>
                              <Button 
                                size="icon" 
                                variant="ghost" 
                                className="h-9 w-9 rounded-xl hover:bg-destructive/5 text-destructive/40 hover:text-destructive"
                                onClick={() => handleDeletePltPlhRecord(record.id, 'PLH')}
                              >
                                <Trash2 className="h-4 w-4" />
                              </Button>
                            </div>
                          </TableCell>
                        </motion.tr>
                      ))
                    ) : (
                      <TableRow className="hover:bg-transparent">
                        <TableCell colSpan={5} className="h-32 text-center font-black text-[10px] uppercase tracking-[0.2em] text-blue-500/20">
                          Tidak ada data penugasan PLH ditemukan.
                        </TableCell>
                      </TableRow>
                    )}
                  </TableBody>
                </Table>
              </div>
            </motion.div>
          </div>
        ) : (
          <div className="flex flex-col items-center justify-center py-20 bg-white rounded-2xl border border-dashed shadow-sm">
            <div className="h-20 w-20 rounded-full bg-primary/5 flex items-center justify-center mb-6">
              <Archive className="h-10 w-10 text-primary/40" />
            </div>
            <h3 className="text-xl font-bold text-[#1A4A9A]">Arsip Dokumen BSKJI</h3>
            <p className="text-sm text-muted-foreground/60 max-w-md text-center mt-2 px-6">
              Akses berkas digital dan dokumen arsip BSKJI melalui portal Google Drive resmi. 
              Pastikan Anda menggunakan akun email resmi kementerian untuk dapat mengakses folder.
            </p>

            <div className="mt-12 grid grid-cols-1 sm:grid-cols-3 gap-6 w-full max-w-5xl px-6">
              {[
                { 
                  title: "SKP Digital", 
                  desc: "Arsip berkas SKP per Tahun", 
                  icon: FileText,
                  subItems: [
                    { label: "SKP Tahun 2024", link: "https://drive.google.com/drive/folders/1AUbnsuuGburDUOmW2EjN1q2HBxNZGC4B?usp=drive_link" },
                    { label: "SKP Tahun 2025", link: "https://drive.google.com/drive/folders/1dxmSz4Hg2r9bi70Sr83HZt9nrkCmSUfS?usp=drive_link" },
                    { label: "SKP Tahun 2026", link: "https://drive.google.com/drive/folders/1NrZwNme-meV5w4dQoo-PlDiEfBuijFdH?usp=drive_link" },
                    { label: "SKP Tahun 2027", link: "#" },
                    { label: "SKP Tahun 2028", link: "#" },
                    { label: "SKP Tahun 2029", link: "#" },
                    { label: "SKP Tahun 2030", link: "#" },
                  ]
                },
                { 
                  title: "Surat Perintah", 
                  desc: "Arsip berkas PLT & PLH per Tahun", 
                  icon: FileCheck,
                  subItems: [
                    { label: "SP Tahun 2024", link: "#" },
                    { label: "SP Tahun 2025", link: "https://drive.google.com/drive/folders/1_Pw651x0UwAbpZsu-YNMoqm8t74j5Qmr?usp=drive_link" },
                    { label: "SP Tahun 2026", link: "https://drive.google.com/drive/folders/1A2zdCJJ0W9BOsFUD9XxY8B8e_j-5sIOU?usp=drive_link" },
                    { label: "SP Tahun 2027", link: "#" },
                    { label: "SP Tahun 2028", link: "#" },
                    { label: "SP Tahun 2029", link: "#" },
                    { label: "SP Tahun 2030", link: "#" },
                  ]
                },
                { title: "Laporan Kinerja", desc: "Laporan periodik unit kerja", icon: LayoutDashboard },
              ].map((item, i) => (
                <div 
                  key={i} 
                  className={cn(
                    "p-6 rounded-2xl border bg-white flex flex-col items-center text-center shadow-sm",
                    "hover:shadow-md transition-shadow"
                  )}
                >
                  <div className="h-12 w-12 rounded-xl bg-primary/5 flex items-center justify-center mb-4">
                    <item.icon className="h-6 w-6 text-primary" />
                  </div>
                  <h4 className="text-sm font-bold text-primary uppercase tracking-wider">{item.title}</h4>
                  <p className="text-[10px] text-muted-foreground mt-1 mb-4">{item.desc}</p>
                  
                  {item.subItems ? (
                    <div className="w-full space-y-2 mt-2">
                      {item.subItems.map((sub, idx) => (
                        <Button 
                          key={idx}
                          variant="outline" 
                          size="sm" 
                          className="w-full justify-start h-8 text-[10px] font-medium bg-muted/5 border-muted-foreground/10 hover:bg-primary/5 hover:text-primary hover:border-primary/20"
                          disabled={!sub.link || sub.link === '#'}
                          onClick={() => sub.link && sub.link !== '#' && window.open(sub.link, '_blank')}
                        >
                          <ExternalLink className="h-3 w-3 mr-2 opacity-50" />
                          {sub.label}
                        </Button>
                      ))}
                    </div>
                  ) : (
                    <Button 
                      variant="ghost" 
                      size="sm" 
                      className="mt-2 text-[10px] font-bold text-muted-foreground/60"
                      disabled
                    >
                      Segera Hadir
                    </Button>
                  )}
                </div>
              ))}
            </div>
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

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
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
                    <Label htmlFor="edit-tanggalSurat">Tanggal Surat</Label>
                    <Input 
                      id="edit-tanggalSurat" 
                      name="tanggalSurat" 
                      type="date"
                      defaultValue={editingPltRecord?.tanggalSurat || editingPlhRecord?.tanggalSurat} 
                      required 
                    />
                  </div>
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

      {/* Edit Unit SK Dialog */}
      <Dialog open={isEditingUnitSK} onOpenChange={setIsEditingUnitSK}>
        <DialogContent className="sm:max-w-[425px]">
          <form onSubmit={handleUpdateUnitSK}>
            <DialogHeader>
              <DialogTitle>Update Informasi SK Unit</DialogTitle>
              <DialogDescription>
                Informasi SK untuk Unit Kerja: {selectedUnitForTeams}
              </DialogDescription>
            </DialogHeader>
            {(() => {
              const unitInfo = unitInfoList.find(u => u.unitKerja === selectedUnitForTeams);
              return (
                <div className="grid gap-4 py-4">
                  <div className="grid gap-2">
                    <Label htmlFor="unit-nomorSK">Nomor SK</Label>
                    <Input id="unit-nomorSK" name="nomorSK" defaultValue={unitInfo?.nomorSK} placeholder="Nomor Surat Keputusan" />
                  </div>
                  <div className="grid gap-2">
                    <Label htmlFor="unit-tanggalSK">Tanggal SK</Label>
                    <Input id="unit-tanggalSK" name="tanggalSK" type="date" defaultValue={unitInfo?.tanggalSK} />
                  </div>
                </div>
              );
            })()}
            <DialogFooter>
              <Button type="button" variant="outline" onClick={() => setIsEditingUnitSK(false)}>Batal</Button>
              <Button type="submit">Simpan Info SK</Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>

      {/* Add Org Node Dialog */}
      <Dialog open={isAddOrgNodeDialogOpen} onOpenChange={setIsAddOrgNodeDialogOpen}>
        <DialogContent className="sm:max-w-[425px]">
          <form onSubmit={handleAddOrgNode}>
            <DialogHeader>
              <DialogTitle>Tambah Jabatan Struktural</DialogTitle>
              <DialogDescription>
                {parentOrgNodeId 
                  ? `Tambahkan sub-jabatan di bawah ${orgNodes.find(n => n.id === parentOrgNodeId)?.namaJabatan}`
                  : 'Tambahkan jabatan tingkat tertinggi (Puncak Struktur)'}
              </DialogDescription>
            </DialogHeader>
            <div className="grid gap-4 py-4">
              <div className="grid gap-2">
                <Label htmlFor="org-namaJabatan">Nama Jabatan</Label>
                <Input id="org-namaJabatan" name="namaJabatan" placeholder="Contoh: Kepala Balai" required />
              </div>
              <div className="grid gap-2">
                <Label htmlFor="org-namaPegawai">Nama Pejabat</Label>
                <Input id="org-namaPegawai" name="namaPegawai" placeholder="Nama lengkap pejabat" required />
              </div>
              <div className="grid gap-2">
                <Label htmlFor="org-nip">NIP</Label>
                <Input id="org-nip" name="nip" placeholder="Nomor Induk Pegawai" />
              </div>
            </div>
            <DialogFooter>
              <Button type="button" variant="outline" onClick={() => setIsAddOrgNodeDialogOpen(false)}>Batal</Button>
              <Button type="submit">Simpan Jabatan</Button>
            </DialogFooter>
          </form>
        </DialogContent>
      </Dialog>

      {/* Edit Org Node Dialog */}
      <Dialog open={isEditOrgNodeDialogOpen} onOpenChange={setIsEditOrgNodeDialogOpen}>
        <DialogContent className="sm:max-w-[425px]">
          <form onSubmit={handleUpdateOrgNode}>
            <DialogHeader>
              <DialogTitle>Edit Jabatan Struktural</DialogTitle>
              <DialogDescription>
                Perbarui informasi jabatan struktural formal
              </DialogDescription>
            </DialogHeader>
            {editingOrgNode && (
              <div className="grid gap-4 py-4">
                <div className="grid gap-2">
                  <Label htmlFor="edit-org-namaJabatan">Nama Jabatan</Label>
                  <Input id="edit-org-namaJabatan" name="namaJabatan" defaultValue={editingOrgNode.namaJabatan} required />
                </div>
                <div className="grid gap-2">
                  <Label htmlFor="edit-org-namaPegawai">Nama Pejabat</Label>
                  <Input id="edit-org-namaPegawai" name="namaPegawai" defaultValue={editingOrgNode.namaPegawai} required />
                </div>
                <div className="grid gap-2">
                  <Label htmlFor="edit-org-nip">NIP</Label>
                  <Input id="edit-org-nip" name="nip" defaultValue={editingOrgNode.nip} />
                </div>
              </div>
            )}
            <DialogFooter>
              <Button type="button" variant="outline" onClick={() => setIsEditOrgNodeDialogOpen(false)}>Batal</Button>
              <Button type="submit">Simpan Perubahan</Button>
            </DialogFooter>
          </form>
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
        <p>&copy; {new Date().getFullYear()} Monitoring Berkas SKP Pegawai. Sistem Monitoring Berkas Kepegawaian.</p>
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
          <FileCheck className="h-5 w-5" />
          <span className="text-[10px] font-bold">Berkas</span>
        </Button>
        <Button 
          variant="ghost" 
          className={cn("flex flex-col items-center gap-1 h-auto py-1 px-0 flex-1", activeTab === 'skp-monitoring' ? "text-primary" : "text-muted-foreground")}
          onClick={() => {
            setActiveTab('skp-monitoring');
            setMonitoringSelectedEmployee(null);
          }}
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
        <Button 
          variant="ghost" 
          className={cn("flex flex-col items-center gap-1 h-auto py-1 px-0 flex-1", activeTab === 'arsip' ? "text-primary" : "text-muted-foreground")}
          onClick={() => setActiveTab('arsip')}
        >
          <Archive className="h-5 w-5" />
          <span className="text-[10px] font-bold">Arsip</span>
        </Button>
      </div>
    </div>
  </div>
);
}
