import React, { useState } from 'react';
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.bundle.min.js';
import * as XLSX from 'xlsx';
import './App.css';

interface Module {
  number: string;
  hours: number;
  status?: 'NC' | 'EF';
}

interface Subject {
  name: string;
  modules: Module[];
  totalHours: number;
}

interface Student {
  name: string;
  subjects: Subject[];
  totalHours: number;
}

interface ExcelData {
  students: Student[];
  summary: {
    totalStudents: number;
    totalStudentsWithHours: number;
    totalSubjects: number;
    totalHours: number;
  };
}

function App() {
  const [hoursData, setHoursData] = useState<ExcelData | null>(null);
  const [modulesData, setModulesData] = useState<ExcelData | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [showTotalsModal, setShowTotalsModal] = useState(false);
  const [hiddenStudents, setHiddenStudents] = useState<Set<string>>(new Set());

  const processExcelFile = (data: ArrayBuffer, isHoursFile: boolean) => {
    try {
      const workbook = XLSX.read(data, { type: 'array' });
      
      if (workbook.SheetNames.length < 3) {
        throw new Error('O arquivo deve conter pelo menos 3 folhas');
      }

      const thirdSheet = workbook.Sheets[workbook.SheetNames[2]];
      const jsonData = XLSX.utils.sheet_to_json(thirdSheet, { header: 1 }) as any[][];

      console.log('Raw Excel Data:', jsonData);

      if (jsonData.length === 0) {
        throw new Error('A terceira folha está vazia');
      }

      const allStudents: Student[] = [];
      let currentRow = isHoursFile ? 3 : 4; // B4 para horas, B5 para módulos
      
      if (isHoursFile) {
        // Processamento do arquivo de horas (mantém o código original)
        // Encontra as disciplinas (linha 1, a partir de D1)
        const subjects = jsonData[0].slice(3).map((subject: string, index: number) => {
          console.log(`Found subject at index ${index}: ${subject}`);
          return {
            name: subject,
            modules: [] as Module[],
            totalHours: 0
          };
        });

        // Encontra os módulos (linha 3, a partir de D3)
        const modules = jsonData[2].slice(3).map((module: string, index: number) => {
          console.log(`Found module at index ${index}: ${module} for subject ${subjects[index].name}`);
          return module;
        });

        // Processa cada aluno
        while (currentRow < jsonData.length && jsonData[currentRow][1]) {
          const studentName = jsonData[currentRow][1] as string;
          console.log(`\nProcessing student: ${studentName}`);
          const studentSubjects: Subject[] = [];

          // Processa cada disciplina para o aluno
          for (let i = 0; i < subjects.length; i++) {
            const subjectModules: Module[] = [];
            let subjectTotalHours = 0;

            // Encontra o índice da próxima disciplina diferente
            let nextSubjectIndex = i + 1;
            while (nextSubjectIndex < subjects.length && subjects[nextSubjectIndex].name === subjects[i].name) {
              nextSubjectIndex++;
            }

            // Processa todos os módulos até a próxima disciplina
            for (let moduleIndex = i; moduleIndex < nextSubjectIndex; moduleIndex++) {
              const value = jsonData[currentRow][moduleIndex + 3];
              const hours = Number(value) || 0;
              if (hours > 0) {
                subjectModules.push({
                  number: modules[moduleIndex],
                  hours: hours
                });
                subjectTotalHours += hours;
              }
            }

            if (subjectModules.length > 0) {
              studentSubjects.push({
                name: subjects[i].name,
                modules: subjectModules,
                totalHours: subjectTotalHours
              });
            }

            i = nextSubjectIndex - 1;
          }

          const studentTotalHours = studentSubjects.reduce((sum, subject) => sum + subject.totalHours, 0);

          if (studentSubjects.length > 0) {
            allStudents.push({
              name: studentName,
              subjects: studentSubjects,
              totalHours: studentTotalHours
            });
          }

          currentRow++;
        }
      } else {
        // Processamento do arquivo de módulos em atraso
        console.log('Processing modules file - Starting at row:', currentRow);
        
        // Encontra as disciplinas (linha 2, a partir de D2)
        const disciplineRow = jsonData[1].slice(3);
        console.log('Discipline row:', disciplineRow);

        // Encontra os módulos (linha 4, a partir de D4)
        const modules = jsonData[3].slice(3);
        console.log('Modules row:', modules);

        // Processa cada aluno a partir de B5
        currentRow = 4; // Começa em B5
        while (currentRow < jsonData.length && jsonData[currentRow][1]) {
          const studentName = jsonData[currentRow][1] as string;
          console.log('\nProcessing student:', studentName);
          
          // Mapa para agrupar módulos por disciplina
          const subjectModulesMap = new Map<string, Module[]>();

          // Processa cada valor na linha do aluno
          for (let colIndex = 3; colIndex < jsonData[currentRow].length; colIndex++) {
            const value = jsonData[currentRow][colIndex];
            if (value === 'NC' || value === 'EF') {
              const moduleIndex = colIndex - 3;
              const currentModule = modules[moduleIndex];
              const discipline = disciplineRow[moduleIndex];

              console.log(`Found module: Discipline=${discipline}, Module=${currentModule}, Status=${value}`);

              if (!subjectModulesMap.has(discipline)) {
                subjectModulesMap.set(discipline, []);
              }

              subjectModulesMap.get(discipline)!.push({
                number: currentModule,
                hours: 0,
                status: value as 'NC' | 'EF'
              });
            }
          }

          // Converte o mapa em array de subjects
          const studentSubjects: Subject[] = Array.from(subjectModulesMap.entries()).map(([name, modules]) => ({
            name,
            modules: modules.sort((a, b) => parseInt(a.number) - parseInt(b.number)),
            totalHours: 0
          }));

          if (studentSubjects.length > 0) {
            console.log('Final subjects for student:', studentSubjects);
            allStudents.push({
              name: studentName,
              subjects: studentSubjects,
              totalHours: 0
            });
          }

          currentRow++;
        }
      }

      const result = {
        students: allStudents,
        summary: {
          totalStudents: allStudents.length,
          totalStudentsWithHours: allStudents.length,
          totalSubjects: allStudents.reduce((max, student) => 
            Math.max(max, student.subjects.length), 0),
          totalHours: allStudents.reduce((sum, student) => sum + student.totalHours, 0)
        }
      };

      console.log('Final processed data:', result);
      return result;
    } catch (err) {
      console.error('Error processing file:', err);
      throw err;
    }
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>, isHoursFile: boolean) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setError(null);
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target?.result as ArrayBuffer;
        if (data) {
          const processedData = processExcelFile(data, isHoursFile);
          if (isHoursFile) {
            setHoursData(processedData);
          } else {
            setModulesData(processedData);
          }
        }
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Erro ao processar o arquivo Excel.');
        console.error(err);
      }
    };

    reader.onerror = () => {
      setError('Erro ao ler o arquivo. Por favor, tente novamente.');
    };

    reader.readAsArrayBuffer(file);
  };

  // Combina os dados dos dois arquivos
  const getCombinedData = () => {
    if (!hoursData || !modulesData) return null;

    console.log('Combining data:');
    console.log('Hours data:', hoursData);
    console.log('Modules data:', modulesData);

    // Cria um mapa de todos os alunos únicos
    const allStudentNames = new Set([
      ...hoursData.students.map(s => s.name),
      ...modulesData.students.map(s => s.name)
    ]);

    const combinedStudents = Array.from(allStudentNames).map(studentName => {
      const hoursStudent = hoursData.students.find(s => s.name === studentName);
      const modulesStudent = modulesData.students.find(s => s.name === studentName);

      // Cria um mapa de todas as disciplinas únicas
      const allSubjectNames = new Set([
        ...(hoursStudent?.subjects || []).map(s => s.name),
        ...(modulesStudent?.subjects || []).map(s => s.name)
      ]);

      // Combina as disciplinas
      const combinedSubjects = Array.from(allSubjectNames).map(subjectName => {
        const hoursSubject = hoursStudent?.subjects.find(s => s.name === subjectName);
        const modulesSubject = modulesStudent?.subjects.find(s => s.name === subjectName);

        // Combina os módulos da disciplina
        const allModules: Module[] = [];
        
        // Adiciona módulos com horas
        if (hoursSubject) {
          hoursSubject.modules.forEach(module => {
            allModules.push({
              number: module.number,
              hours: module.hours,
              status: undefined
            });
          });
        }

        // Adiciona ou atualiza módulos em atraso
        if (modulesSubject) {
          modulesSubject.modules.forEach(module => {
            const existingModule = allModules.find(m => m.number === module.number);
            if (existingModule) {
              existingModule.status = module.status;
            } else {
              allModules.push({
                number: module.number,
                hours: 0,
                status: module.status
              });
            }
          });
        }

        // Ordena os módulos por número
        allModules.sort((a, b) => {
          const numA = parseInt(a.number);
          const numB = parseInt(b.number);
          return numA - numB;
        });

        return {
          name: subjectName,
          modules: allModules,
          totalHours: hoursSubject?.totalHours || 0
        };
      });

      return {
        name: studentName,
        subjects: combinedSubjects.sort((a, b) => a.name.localeCompare(b.name)),
        totalHours: hoursStudent?.totalHours || 0
      };
    });

    const result = {
      students: combinedStudents.sort((a, b) => a.name.localeCompare(b.name)),
      summary: {
        totalStudents: combinedStudents.length,
        totalStudentsWithHours: combinedStudents.filter(s => s.totalHours > 0).length,
        totalSubjects: Math.max(
          hoursData.summary.totalSubjects,
          modulesData.summary.totalSubjects
        ),
        totalHours: combinedStudents.reduce((sum, student) => sum + student.totalHours, 0)
      }
    };

    console.log('Final combined data:', result);
    return result;
  };

  const combinedData = getCombinedData();

  // Adiciona função para remover aluno
  const handleRemoveStudent = (studentName: string) => {
    setHiddenStudents(prev => {
      const newSet = new Set(prev);
      newSet.add(studentName);
      return newSet;
    });
  };

  return (
    <div className="App py-4">
      <div className="container">
        <div className="row justify-content-center">
          <div className="col-lg-8 text-center">
            <div className="d-flex align-items-center justify-content-between mb-4">
              <h1 className="display-5 fw-bold text-white mb-0">
                <i className="bi bi-file-earmark-excel me-2"></i>
                Processador de PRAs
              </h1>
              {combinedData && combinedData.students.length > 0 && (
                <button 
                  className="btn btn-info"
                  onClick={() => setShowTotalsModal(true)}
                >
                  <i className="bi bi-table me-2"></i>
                  Ver Totais
                </button>
              )}
            </div>
            
            <div className="card shadow-lg border-0 rounded-3 mb-5">
              <div className="card-header bg-dark text-white py-3">
                <h5 className="card-title mb-0">
                  <i className="bi bi-upload me-2"></i>
                  Upload dos Arquivos Excel
                </h5>
              </div>
              <div className="card-body">
                <div className="row g-3">
                  <div className="col-md-6">
                    <div className="upload-container p-3 rounded-3 bg-dark bg-opacity-10">
                      <label className="form-label text-white-50">
                        <i className="bi bi-clock-history me-2"></i>
                        Arquivo de Horas de Reposição
                      </label>
                      <input
                        type="file"
                        className="form-control form-control-sm"
                        accept=".xlsx, .xls"
                        onChange={(e) => handleFileUpload(e, true)}
                      />
                    </div>
                  </div>
                  <div className="col-md-6">
                    <div className="upload-container p-3 rounded-3 bg-dark bg-opacity-10">
                      <label className="form-label text-white-50">
                        <i className="bi bi-exclamation-triangle me-2"></i>
                        Arquivo de Módulos em Atraso
                      </label>
                      <input
                        type="file"
                        className="form-control form-control-sm"
                        accept=".xlsx, .xls"
                        onChange={(e) => handleFileUpload(e, false)}
                      />
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {error && (
          <div className="row justify-content-center mb-4">
            <div className="col-lg-8">
              <div className="alert alert-danger d-flex align-items-center" role="alert">
                <i className="bi bi-exclamation-triangle-fill me-2"></i>
                <div>{error}</div>
              </div>
            </div>
          </div>
        )}

        {combinedData && combinedData.students.map((student, studentIndex) => (
          <div key={studentIndex} className="student-table mb-5">
            <div className="d-flex align-items-center mb-3 bg-dark bg-opacity-50 p-3 rounded-3">
              <i className="bi bi-person-circle fs-4 me-2 text-white-50"></i>
              <h3 className="mb-0">Aluno: {student.name}</h3>
              <div className="ms-auto d-flex gap-4">
                {student.subjects.some(subject => subject.modules.some(m => m.hours > 0)) && (
                  <div className="text-warning">
                    <i className="bi bi-clock-history me-2"></i>
                    {student.subjects.reduce((total, subject) => 
                      total + subject.modules.reduce((sum, module) => sum + module.hours, 0), 0)}h a repor
                  </div>
                )}
                {student.subjects.some(subject => subject.modules.some(m => m.status)) && (
                  <div className="text-danger">
                    <i className="bi bi-exclamation-triangle me-2"></i>
                    {student.subjects.reduce((total, subject) => 
                      total + subject.modules.filter(m => m.status).length, 0)} módulos em atraso
                  </div>
                )}
              </div>
            </div>

            {/* Tabela de Horas de Reposição */}
            {student.subjects.some(subject => subject.modules.some(m => m.hours > 0)) && (
              <div className="table-responsive shadow-sm rounded-3 mb-4">
                <table className="table table-bordered mb-0">
                  <thead className="table-primary">
                    <tr>
                      <th colSpan={3} className="text-center">REPOSIÇÃO DE HORAS</th>
                    </tr>
                    <tr>
                      <th>DISCIPLINA</th>
                      <th>Horas por repor</th>
                      <th>TAREFAS</th>
                    </tr>
                  </thead>
                  <tbody>
                    {student.subjects
                      .filter(subject => subject.modules.some(m => m.hours > 0))
                      .sort((a, b) => a.name.localeCompare(b.name))
                      .flatMap(subject => {
                        // Pega todos os módulos com horas > 0
                        const modulesWithHours = subject.modules
                          .filter(module => module.hours > 0)
                          .sort((a, b) => parseInt(a.number) - parseInt(b.number));

                        // Cria uma linha para cada módulo
                        return modulesWithHours.map((module, index) => (
                          <tr key={`${subject.name}-hours-${module.number}`}>
                            {index === 0 ? <td>{subject.name}</td> : <td style={{ border: 0 }}></td>}
                            <td>{`M${module.number} - ${module.hours}h`}</td>
                            <td></td>
                          </tr>
                        ));
                      })}
                  </tbody>
                </table>
              </div>
            )}

            {/* Tabela de Módulos em Atraso */}
            {student.subjects.some(subject => subject.modules.some(m => m.status)) && (
              <div className="table-responsive shadow-sm rounded-3">
                <table className="table table-bordered mb-0">
                  <thead className="table-primary">
                    <tr>
                      <th colSpan={3} className="text-center">RECUPERAÇÃO DE MÓDULOS EM ATRASO</th>
                    </tr>
                    <tr>
                      <th>DISCIPLINA</th>
                      <th>Módulo</th>
                      <th>TAREFAS</th>
                    </tr>
                  </thead>
                  <tbody>
                    {student.subjects
                      .filter(subject => subject.modules.some(m => m.status))
                      .sort((a, b) => a.name.localeCompare(b.name))
                      .map(subject => 
                        subject.modules
                          .filter(module => module.status)
                          .sort((a, b) => parseInt(a.number) - parseInt(b.number))
                          .map((module, moduleIndex) => (
                            <tr key={`${subject.name}-arrears-${module.number}`}>
                              {moduleIndex === 0 ? <td>{subject.name}</td> : <td style={{ border: 0 }}></td>}
                              <td>{`M${module.number} - ${module.status}`}</td>
                              <td></td>
                            </tr>
                          ))
                      )}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        ))}

        {/* Modal de Totais */}
        {showTotalsModal && (
          <>
            <div className="modal-backdrop fade show"></div>
            <div 
              className="modal fade show d-block" 
              tabIndex={-1}
              style={{ backgroundColor: 'rgba(0, 0, 0, 0.5)' }}
              onClick={(e) => {
                if (e.target === e.currentTarget) {
                  setShowTotalsModal(false);
                }
              }}
            >
              <div className="modal-dialog modal-lg modal-dialog-centered">
                <div 
                  className="modal-content border-0 shadow"
                  onClick={e => e.stopPropagation()}
                >
                  <div className="modal-header bg-dark text-white border-0">
                    <h5 className="modal-title">
                      <i className="bi bi-table me-2"></i>
                      Totais por Aluno
                    </h5>
                    <button 
                      type="button" 
                      className="btn-close btn-close-white" 
                      onClick={() => setShowTotalsModal(false)}
                      aria-label="Close"
                    ></button>
                  </div>
                  <div className="modal-body p-0">
                    <div className="table-responsive">
                      <table className="table table-hover table-striped mb-0">
                        <thead className="table-dark">
                          <tr>
                            <th className="text-start" style={{ minWidth: '250px' }}>Aluno</th>
                            <th className="text-center" style={{ width: '150px' }}>Horas a Repor</th>
                            <th className="text-center" style={{ width: '150px' }}>Módulos em Atraso</th>
                          </tr>
                        </thead>
                        <tbody>
                          {combinedData?.students
                            .filter(student => !hiddenStudents.has(student.name))
                            .map((student, index) => {
                            const totalHours = student.subjects.reduce((total, subject) => 
                              total + subject.modules.reduce((sum, module) => sum + module.hours, 0), 0);
                            const totalModules = student.subjects.reduce((total, subject) => 
                              total + subject.modules.filter(m => m.status).length, 0);
                            
                            return (
                              <tr key={index}>
                                <td>
                                  <div className="d-flex justify-content-between align-items-center gap-2">
                                    <span className="text-truncate">{student.name}</span>
                                    <button
                                      className="btn btn-danger btn-sm d-flex align-items-center justify-content-center p-1"
                                      onClick={() => handleRemoveStudent(student.name)}
                                      title="Remover da lista"
                                      style={{ width: '24px', height: '24px' }}
                                    >
                                      <i className="bi bi-x-lg"></i>
                                    </button>
                                  </div>
                                </td>
                                <td className="text-center">
                                  {totalHours > 0 ? (
                                    <span className="badge bg-warning text-dark">
                                      {totalHours}h
                                    </span>
                                  ) : '-'}
                                </td>
                                <td className="text-center">
                                  {totalModules > 0 ? (
                                    <span className="badge bg-danger">
                                      {totalModules}
                                    </span>
                                  ) : '-'}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot className="table-dark">
                          <tr>
                            <td>
                              <div className="d-flex justify-content-between align-items-center">
                                <strong>TOTAL</strong>
                                {hiddenStudents.size > 0 && (
                                  <button
                                    className="btn btn-link text-info p-0 ms-2"
                                    onClick={() => setHiddenStudents(new Set())}
                                    title="Restaurar todos os alunos"
                                  >
                                    <i className="bi bi-arrow-counterclockwise"></i>
                                  </button>
                                )}
                              </div>
                            </td>
                            <td className="text-center">
                              <strong>
                                {combinedData?.students
                                  .filter(student => !hiddenStudents.has(student.name))
                                  .reduce((total, student) => 
                                    total + student.subjects.reduce((subTotal, subject) => 
                                      subTotal + subject.modules.reduce((sum, module) => sum + module.hours, 0), 0), 0)}h
                              </strong>
                            </td>
                            <td className="text-center">
                              <strong>
                                {combinedData?.students
                                  .filter(student => !hiddenStudents.has(student.name))
                                  .reduce((total, student) => 
                                    total + student.subjects.reduce((subTotal, subject) => 
                                      subTotal + subject.modules.filter(m => m.status).length, 0), 0)}
                              </strong>
                            </td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                  <div className="modal-footer border-0">
                    <button 
                      type="button" 
                      className="btn btn-secondary" 
                      onClick={() => setShowTotalsModal(false)}
                    >
                      Fechar
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

export default App;
