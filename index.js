const mysql = require("mysql"); // PLUGIN DE CONEXION
const util = require("util"); //PLUGIN DE MANEJO DE CONEXION
const moment = require("moment"); //Manejo de tiempo
var xl = require("excel4node"); //EXCEL
var archiver = require("archiver"); //COMPRIMIR
var fs = require("fs");
const mysqlssh = require("mysql-ssh");

// CONFIGURACION DB
// const conn = mysql.createConnection({
//   host: "162.214.88.238",
//   user: "wwaula_cham2",
//   password: "P.ptDcgPHlEnQAP9ZBb31",
//   database: "wwaula_cham2",
// });

const secondsFix = (time) => {
  try {
    var hours = Math.floor(time / 3600);
    var minutes = Math.floor((time - hours * 3600) / 60);
    var seconds = time - hours * 3600 - minutes * 60;

    return `${hours < 9 ? `0${hours}` : hours}:${
      minutes < 9 ? `0${minutes}` : minutes
    }:${seconds < 9 ? `0${seconds}` : seconds}`;
  } catch (error) {
    return "00:00:00";
  }
};

const timeFix = (time) => {
  try {
    const h = time.split(":")[0];
    const m = time.split(":")[1];
    const s = time.split(":")[2];

    return parseInt(s) + parseInt(m) * 60 + parseInt(h) * 3600;
  } catch (error) {
    return 0;
  }
};

const convertAllSessions = (list) => {
  try {
    let time = 0;

    list.forEach((item) => {
      time =
        moment(item.logout_course_date).diff(
          moment(item.login_course_date),
          "seconds"
        ) + time;
    });
    return time;
  } catch (error) {
    return 0;
  }
};

const getBackground = (time) => {
  try {
    const currentTime = timeFix(time);

    if (currentTime <= 1)
      return {
        fill: {
          type: "pattern",
          patternType: "solid",
          bgColor: "#EF9A9A",
          fgColor: "#EF9A9A",
        },
      };
    else if (currentTime >= 2 && currentTime <= 1800)
      return {
        fill: {
          type: "pattern",
          patternType: "solid",
          bgColor: "#FFD1FF",
          fgColor: "#FFD1FF",
        },
      };
    else if (currentTime >= 1801 && currentTime <= 3599)
      return {
        fill: {
          type: "pattern",
          patternType: "solid",
          bgColor: "#FFFFCC",
          fgColor: "#FFFFCC",
        },
      };
    else
      return {
        fill: {
          type: "pattern",
          patternType: "solid",
          bgColor: "#64FFDA",
          fgColor: "#64FFDA",
        },
      };
  } catch (error) {
    return {};
  }
};

const SCHOOL = "d121";
mysqlssh
  .connect(
    // {
    //   host: "162.214.88.238",
    //   port: 22022,
    //   user: "root",
    //   password: "Flatrone194166#",
    // },
    // {
    //   host: "162.214.88.238",
    //   // user: "wwaula_cham2", //huan
    //   // password: "P.ptDcgPHlEnQAP9ZBb31",
    //   // database: "wwaula_cham2",
    //   user: "aulacodi_cham2",
    //   password: "Q.vsS0IGj73TOw0jZOp57",
    //   database: "aulacodi_cham2",
    // }
    {
      host: "162.214.97.148",
      user: "root",
      password: "flatrone194166#",
    },
    {
      host: "162.214.97.148",
      user: "admin_ad129", //codiple1_generica
      password: "h?k}I8$vStY]",
      database: "admin_ad129",
    }
  )
  .then(async (client) => {
    const query = util.promisify(client.query).bind(client);

    try {
      let finalData = [];

      // QUERIES
      const category = await query(
        "SELECT id, name, code,tree_pos, parent_id FROM course_category "
      );
      const cursos = await query(
        "SELECT id, title,  category_code FROM course"
      );
      const inscripciones = await query(
        "SELECT u.firstname ,u.isTeacher , u.lastname ,u.id, u.username ,cru.c_id ,u.status FROM course_rel_user cru join user u on u.id = cru.user_id "
      ); //status 5: alumno , 1:docente , 4:recursos humanos

      const timeInCourse = await query(
        "SELECT  c_id, user_id, login_course_date, logout_course_date, counter  FROM track_e_course_access "
      );
      const ListOfAccess = await query(
        "SELECT login_course_date,logout_course_date ,user_id FROM track_e_course_access"
      );
      const listOfUser = await query(
        "SELECT id, firstname , lastname ,status,isTeacher FROM `user`"
      );
      const lecciones = await query("SELECT  c_id FROM c_lp ");
      const leccionesUser = await query(
        "SELECT c_id,  user_id, progress FROM c_lp_view       "
      );
      const ejercicios = await query(
        "SELECT  c_id  FROM c_quiz where active <> -1"
      );
      const meetings = await query("SELECT c_id FROM plugin_bbb_meeting  ");
      const documentos = await query(
        "SELECT path , c_id FROM c_document WHERE  filetype  = 'file'"
      );
      const tareas = await query(`
      SELECT 
      DISTINCT parent_id,
      c_id
      FROM  c_student_publication csp `);

      finalData = category
        .map((cat) => {
          return {
            ...cat,
            cursos: cursos
              .filter(
                (c) => c.category_code.toLowerCase() == cat.code.toLowerCase()
              )
              .map((c) => {
                return {
                  ...c,
                  inscripciones: inscripciones
                    .filter((i) => i.c_id == c.id)

                    .map((i) => {
                      let time = 0;
                      let ingresos = timeInCourse.filter(
                        (tic) => tic.c_id == c.id && tic.user_id == i.id
                      );

                      ingresos.forEach((item) => {
                        time =
                          time +
                          moment(item.logout_course_date).diff(
                            moment(item.login_course_date),
                            "seconds"
                          );
                      });

                      return {
                        ...i,
                        sesiones: secondsFix(time),
                      };
                    }),
                };
              }),
          };
        })
        .filter((item) => item.cursos.length > 0);

      let dataFormated = finalData.map((curso) => {
        let currentAlumnos = [];

        curso.cursos.forEach((asig) => {
          asig.inscripciones.forEach((ins) => {
            const currenAl = currentAlumnos.findIndex(
              (itemCu) => itemCu.id == ins.id
            );

            if (currenAl == -1) {
              currentAlumnos.push({
                firstname: ins.firstname,
                lastname: ins.lastname,
                id: ins.id,
                username: ins.username,
                c_id: ins.c_id,
                status: ins.status,
                isTeacher: ins.isTeacher,
                asignaturas: [{ c_id: asig.id, time: ins.sesiones }],
              });
            } else {
              currentAlumnos[currenAl] = {
                ...currentAlumnos[currenAl],
                asignaturas: [
                  ...currentAlumnos[currenAl].asignaturas,
                  { c_id: asig.id, time: ins.sesiones },
                ],
              };
            }
          });
        });

        return {
          ...curso,

          alumnos: currentAlumnos,
          cursos: curso.cursos.map((asig) => ({
            id: asig.id,
            title: asig.title,
            category_code: asig.category_code,
            inscripciones: inscripciones.filter(
              (item) => item.c_id == asig.id && item.status == 1
            ),
            lecciones: lecciones.filter((item) => item.c_id == asig.id),
            ejercicios: ejercicios.filter((item) => item.c_id == asig.id),
            meetings: meetings.filter((item) => item.c_id == asig.id),
            documentos: documentos
              .filter((item) => item.c_id == asig.id)
              .filter((item) => {
                let search = item.path.search("/borrar/");

                if (search != -1) return false;
                search = item.path.search("/certificates/");

                if (search != -1) return false;

                return true;
              }),
            tareas: tareas.filter((item) => item.c_id == asig.id),
          })),
        };
      });

      // ARCHIVOS EXTERNOS
      // TIEMPO TOTAL
      const fileTotalDocentes = async () => {
        let wb = new xl.Workbook({});
        // ESTILOS
        // Create a reusable style
        var header = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          fill: {
            type: "pattern",
            patternType: "solid",
            bgColor: "#0277bd",
            fgColor: "#0277bd",
          },
          font: {
            bold: true,
            color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            // horizontal: "center",
          },
        });
        var asignaturaStyle = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          fill: {
            type: "pattern",
            patternType: "solid",
            bgColor: "#D9E1F2",
            fgColor: "#D9E1F2",
          },
          font: {
            bold: true,
            // color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            horizontal: "center",
            wrapText: true,
          },
        });
        var alumnoStyle1 = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          fill: {
            type: "pattern",
            patternType: "solid",
            bgColor: "#D9D9D9",
            fgColor: "#D9D9D9",
          },
          font: {
            bold: true,
            // color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            // horizontal: "center",
            wrapText: true,
          },
        });
        var alumnoStyle2 = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },

          font: {
            bold: true,
            // color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            // horizontal: "center",
            wrapText: true,
          },
        });
        var valorStyle = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          // fill: {
          //   type: "pattern",
          //   patternType: "solid",
          //   bgColor: "#4fc3f7",
          //   fgColor: "#4fc3f7",
          // },
          font: {
            // color: "#ffffff",
            size: 12,
          },
          alignment: {
            horizontal: "center",
            wrapText: true,
          },
        });

        const crearhoja = (name, type) => {
          // Add Worksheets to the workbook
          let ws = wb.addWorksheet(name);

          // HEADERS
          let BASE_COL = 1;
          ws.row(1).setHeight(30);
          ws.cell(1, BASE_COL, 1, 4, true).string(name).style(header);
          ws.column(1).setWidth(5);
          ws.column(2).setWidth(25);
          ws.column(3).setWidth(25);
          ws.column(4).setWidth(20);
          ws.column(5).setWidth(20);
          ws.cell(2, BASE_COL).string("Nro").style(asignaturaStyle);
          ws.cell(2, BASE_COL + 1)
            .string("Apellidos")
            .style(asignaturaStyle);
          ws.cell(2, BASE_COL + 2)
            .string("Nombres")
            .style(asignaturaStyle);

          ws.cell(2, BASE_COL + 3)
            .string("Tiempo en plataforma")
            .style(asignaturaStyle);

          let docenteList = listOfUser
            .filter((item) => item.status == type)
            .map((item) => {
              const ingresos = ListOfAccess.filter(
                (item2) => item2.user_id == item.id
              );
              return {
                ...item,
                ingresos: ingresos.length,
                conexiones: convertAllSessions(ingresos),
              };
            });

          const BASE_ROW = 3;
          docenteList
            .sort((a, b) =>
              a.conexiones > b.conexiones
                ? 1
                : a.conexiones < b.conexiones
                ? -1
                : 0
            )
            .forEach((docente, i) => {
              ws.cell(BASE_ROW + i, 1).number(i + 1);
              ws.cell(BASE_ROW + i, 2).string(docente.firstname.toUpperCase());
              ws.cell(BASE_ROW + i, 3).string(docente.lastname.toUpperCase());

              ws.cell(BASE_ROW + i, 4)
                .string(secondsFix(docente.conexiones))

                .style({
                  ...valorStyle,
                  ...getBackground(secondsFix(docente.conexiones)),
                  numberFormat: "[h]:mm:ss",
                });
            });
        };

        crearhoja("Docentes", 1);
        crearhoja("RRHH", 4);
        crearhoja("Alumnos", 5);

        // CREACION DE ARCHIVO EXCEL
        const fileCreate = await wb.writeToBuffer();
        fs.writeFileSync(`acceso-plataforma.xlsx`, fileCreate);
      };

      const detalleCurso = async () => {
        let wb = new xl.Workbook({});
        // ESTILOS
        // Create a reusable style
        var header = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          fill: {
            type: "pattern",
            patternType: "solid",
            bgColor: "#0277bd",
            fgColor: "#0277bd",
          },
          font: {
            bold: true,
            color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            // horizontal: "center",
          },
        });
        var asignaturaStyle = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          fill: {
            type: "pattern",
            patternType: "solid",
            bgColor: "#D9E1F2",
            fgColor: "#D9E1F2",
          },
          font: {
            bold: true,
            // color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            horizontal: "center",
            wrapText: true,
          },
        });
        var alumnoStyle1 = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          fill: {
            type: "pattern",
            patternType: "solid",
            bgColor: "#D9D9D9",
            fgColor: "#D9D9D9",
          },
          font: {
            bold: true,
            // color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            // horizontal: "center",
            wrapText: true,
          },
        });
        var alumnoStyle2 = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },

          font: {
            bold: true,
            // color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            // horizontal: "center",
            wrapText: true,
          },
        });
        var valorStyle = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          // fill: {
          //   type: "pattern",
          //   patternType: "solid",
          //   bgColor: "#4fc3f7",
          //   fgColor: "#4fc3f7",
          // },
          font: {
            // color: "#ffffff",
            size: 12,
          },
          alignment: {
            horizontal: "center",
            wrapText: true,
          },
        });

        const getColors = (numeroAsignado) => {
          try {
            if (numeroAsignado < 1)
              return {
                fill: {
                  type: "pattern",
                  patternType: "solid",
                  bgColor: "#EF9A9A",
                  fgColor: "#EF9A9A",
                },
              };
            else if (numeroAsignado >= 1 && numeroAsignado <= 3)
              return {
                fill: {
                  type: "pattern",
                  patternType: "solid",
                  bgColor: "#FFD1FF",
                  fgColor: "#FFD1FF",
                },
              };
            else if (numeroAsignado >= 4 && numeroAsignado <= 6)
              return {
                fill: {
                  type: "pattern",
                  patternType: "solid",
                  bgColor: "#FFFFCC",
                  fgColor: "#FFFFCC",
                },
              };
            else
              return {
                fill: {
                  type: "pattern",
                  patternType: "solid",
                  bgColor: "#64FFDA",
                  fgColor: "#64FFDA",
                },
              };
          } catch (error) {
            return {};
          }
        };

        const crearhoja = (name) => {
          // Add Worksheets to the workbook
          let ws = wb.addWorksheet(name);

          // HEADERS
          let BASE_COL = 1;
          ws.row(1).setHeight(30);
          ws.cell(1, BASE_COL, 1, 7, true).string(name).style(header);
          ws.column(1).setWidth(25);
          ws.column(2).setWidth(25);
          ws.column(3).setWidth(20);
          ws.column(4).setWidth(20);
          ws.column(5).setWidth(20);
          ws.column(6).setWidth(20);
          ws.column(7).setWidth(30);

          ws.cell(3, 1).string("Curso").style(header);
          ws.cell(3, 2).string("Asignatura").style(header);
          ws.cell(3, 3).string("Cantidad de clases").style(header);
          ws.cell(3, 4).string("Cantidad de ejercicios").style(header);
          // ws.cell(3, 5).string("Cantidad de Videoconferencias").style(header);
          ws.cell(3, 5).string("Cantidad de Archivos").style(header);
          ws.cell(3, 6).string("Cantidad de Tareas").style(header);
          ws.cell(3, 7).string("Docente(s)").style(header);

          let BASEASIGANTURAS = 4;
          // ASIGNATURAS
          dataFormated.forEach((asignatura, i) => {
            ws.cell(BASEASIGANTURAS, 1)
              .string(asignatura.name)
              .style(asignaturaStyle);

            // CURSO
            asignatura.cursos.forEach((curso, j) => {
              ws.cell(BASEASIGANTURAS, 2)
                .string(curso.title)
                .style(asignaturaStyle);

              // LECCIONES
              ws.cell(BASEASIGANTURAS, 3)
                .number(curso.lecciones.length)
                .style({ ...valorStyle, ...getColors(curso.lecciones.length) });
              // EJERCICIOS
              ws.cell(BASEASIGANTURAS, 4)
                .number(curso.ejercicios.length)
                .style({
                  ...valorStyle,
                  ...getColors(curso.ejercicios.length),
                });
              // EJERCICIOS
              // ws.cell(BASEASIGANTURAS, 5)
              //   .number(curso.meetings.length)
              //   .style({ ...valorStyle, ...getColors(curso.meetings.length) });
              // Archivos
              ws.cell(BASEASIGANTURAS, 5)
                .number(curso.documentos.length)
                .style({
                  ...valorStyle,
                  ...getColors(curso.documentos.length),
                });
              // Tareas
              ws.cell(BASEASIGANTURAS, 6)
                .number(curso.tareas.length)
                .style({ ...valorStyle, ...getColors(curso.tareas.length) });
              ws.cell(BASEASIGANTURAS, 7)
                .string(
                  `${curso.inscripciones
                    .map((item) => `${item.firstname} ${item.lastname}`)
                    .join(", ")}.`
                )
                .style({ ...valorStyle });

              BASEASIGANTURAS++;
            });
            BASEASIGANTURAS++;
          });

          // ws.cell(2, BASE_COL).string("Nro").style(asignaturaStyle);
          // ws.cell(2, BASE_COL + 1)
          //   .string("Apellidos")
          //   .style(asignaturaStyle);
          // ws.cell(2, BASE_COL + 2)
          //   .string("Nombres")
          //   .style(asignaturaStyle);

          // ws.cell(2, BASE_COL + 3)
          //   .string("Tiempo en plataforma")
          //   .style(asignaturaStyle);

          // let docenteList = listOfUser
          //   .filter((item) => item.status == type)
          //   .map((item) => {
          //     const ingresos = ListOfAccess.filter(
          //       (item2) => item2.user_id == item.id
          //     );
          //     return {
          //       ...item,
          //       ingresos: ingresos.length,
          //       conexiones: convertAllSessions(ingresos),
          //     };
          //   });

          // const BASE_ROW = 3;
          // docenteList
          //   .sort((a, b) =>
          //     a.conexiones > b.conexiones
          //       ? 1
          //       : a.conexiones < b.conexiones
          //       ? -1
          //       : 0
          //   )
          //   .forEach((docente, i) => {
          //     ws.cell(BASE_ROW + i, 1).number(i + 1);
          //     ws.cell(BASE_ROW + i, 2).string(docente.firstname.toUpperCase());
          //     ws.cell(BASE_ROW + i, 3).string(docente.lastname.toUpperCase());

          //     ws.cell(BASE_ROW + i, 4)
          //       .string(secondsFix(docente.conexiones))

          //       .style({
          //         ...valorStyle,
          //         ...getBackground(secondsFix(docente.conexiones)),
          //         numberFormat: "[h]:mm:ss",
          //       });
          // });
        };

        crearhoja("Cursos");

        // CREACION DE ARCHIVO EXCEL
        const fileCreate = await wb.writeToBuffer();
        fs.writeFileSync(`detalle_curso.xlsx`, fileCreate);
      };

      detalleCurso();
      fileTotalDocentes();

      // CREAR CADA EXCEL
      for (let index = 0; index < dataFormated.length; index++) {
        // Create a new instance of a Workbook class
        let wb = new xl.Workbook({});
        // ESTILOS
        // Create a reusable style
        var header = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          fill: {
            type: "pattern",
            patternType: "solid",
            bgColor: "#0277bd",
            fgColor: "#0277bd",
          },
          font: {
            bold: true,
            color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            // horizontal: "center",
          },
        });
        var asignaturaStyle = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          fill: {
            type: "pattern",
            patternType: "solid",
            bgColor: "#D9E1F2",
            fgColor: "#D9E1F2",
          },
          font: {
            bold: true,
            // color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            horizontal: "center",
            wrapText: true,
          },
        });
        var alumnoStyle1 = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          fill: {
            type: "pattern",
            patternType: "solid",
            bgColor: "#D9D9D9",
            fgColor: "#D9D9D9",
          },
          font: {
            bold: true,
            // color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            // horizontal: "center",
            wrapText: true,
          },
        });
        var alumnoStyle2 = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },

          font: {
            bold: true,
            // color: "#ffffff",
            size: 10,
          },
          alignment: {
            vertical: "center",
            // horizontal: "center",
            wrapText: true,
          },
        });
        var valorStyle = wb.createStyle({
          border: {
            // §18.8.4 border (Border)
            left: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            right: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            top: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
            bottom: {
              style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
              color: "#000000", // HTML style hex value
            },
          },
          // fill: {
          //   type: "pattern",
          //   patternType: "solid",
          //   bgColor: "#4fc3f7",
          //   fgColor: "#4fc3f7",
          // },
          font: {
            // color: "#ffffff",
            size: 12,
          },
          alignment: {
            horizontal: "center",
            wrapText: true,
          },
        });

        const fileStudent = async () => {
          // Add Worksheets to the workbook
          let ws = wb.addWorksheet("Alumnos");

          // HEADERS
          let BASE_COL = 1;
          ws.row(1).setHeight(30);
          ws.cell(1, BASE_COL, 1, 3, true).string("ALUMNOS").style(header);
          ws.column(1).setWidth(5);
          ws.column(2).setWidth(25);
          ws.column(3).setWidth(25);
          ws.column(4).setWidth(20);
          ws.column(5).setWidth(20);
          ws.cell(2, BASE_COL).string("Nro").style(asignaturaStyle);
          ws.cell(2, BASE_COL + 1)
            .string("Apellidos")
            .style(asignaturaStyle);
          ws.cell(2, BASE_COL + 2)
            .string("Nombres")
            .style(asignaturaStyle);
          ws.cell(2, BASE_COL + 3)
            .string("Totales / Suma")
            .style(asignaturaStyle);
          ws.cell(2, BASE_COL + 4)
            .string("Promedio")
            .style(asignaturaStyle);

          //RESUMEN
          ws.cell(1, BASE_COL + 3, 1, BASE_COL + 4, true)
            .string("RESUMEN DE HORAS")
            .style(header);
          //ASIGNATURAS
          ws.cell(
            1,
            BASE_COL + 5,
            1,
            BASE_COL + 4 + dataFormated[index].cursos.length,
            true
          )
            .string("ASIGNATURAS")
            .style(header);
          dataFormated[index].cursos.forEach((asignatura, i) => {
            ws.column(BASE_COL + 5 + i).setWidth(15);
            ws.cell(2, BASE_COL + 5 + i)
              .string(asignatura.title)
              .style(asignaturaStyle);
          });

          //   alumnos
          dataFormated[index].alumnos
            .filter(
              (alumno) => alumno.status == 5 //&&
              // alumno.username.search("e") == -1 &&
              // alumno.username.search("E") == -1
            )
            .forEach((alumno, i) => {
              ws.cell(3 + i, 1)
                .number(i + 1)
                .style(i % 2 == 0 ? alumnoStyle1 : alumnoStyle2);
              ws.cell(3 + i, 2)
                .string(alumno.lastname)
                .style(i % 2 == 0 ? alumnoStyle1 : alumnoStyle2);
              ws.cell(3 + i, 3)
                .string(alumno.firstname)
                .style(i % 2 == 0 ? alumnoStyle1 : alumnoStyle2);

              let totalTime = 0;

              // TOTAL DE HORAS
              dataFormated[index].cursos.forEach((asignatura, j) => {
                let currentTime = alumno.asignaturas.find(
                  (item) => item.c_id == asignatura.id
                );

                if (!currentTime) return;

                totalTime = totalTime + timeFix(currentTime.time);
              });

              // TOTAL
              ws.cell(3 + i, 4)
                .string(secondsFix(totalTime))
                .style({
                  ...valorStyle,
                  ...getBackground(secondsFix(totalTime)),
                  numberFormat: "[h]:mm:ss",
                });
              // PROMEDIO
              ws.cell(3 + i, 5)
                .string(
                  secondsFix(parseInt(totalTime / alumno.asignaturas.length))
                )
                .style({
                  ...valorStyle,
                  ...getBackground(
                    secondsFix(parseInt(totalTime / alumno.asignaturas.length))
                  ),
                  numberFormat: "[h]:mm:ss",
                });

              ws.column(5).freeze();
              dataFormated[index].cursos.forEach((asignatura, j) => {
                const currentTime = alumno.asignaturas.find(
                  (item) => item.c_id == asignatura.id
                );

                ws.cell(3 + i, BASE_COL + 5 + j)
                  .string(currentTime ? currentTime.time : "00:00:00")
                  .style({
                    ...valorStyle,
                    ...getBackground(
                      currentTime ? currentTime.time : "00:00:00"
                    ),
                    numberFormat: "[h]:mm:ss",
                  });
              });
            });
        };
        const fileDocente = async () => {
          // Add Worksheets to the workbook
          let ws = wb.addWorksheet("Docentes");

          // HEADERS
          let BASE_COL = 1;
          ws.row(1).setHeight(30);
          ws.cell(1, BASE_COL, 1, 3, true).string("DOCENTES").style(header);
          ws.column(1).setWidth(5);
          ws.column(2).setWidth(25);
          ws.column(3).setWidth(25);
          ws.column(4).setWidth(20);
          ws.column(5).setWidth(20);
          ws.cell(2, BASE_COL).string("Nro").style(asignaturaStyle);
          ws.cell(2, BASE_COL + 1)
            .string("Apellidos")
            .style(asignaturaStyle);
          ws.cell(2, BASE_COL + 2)
            .string("Nombres")
            .style(asignaturaStyle);
          ws.cell(2, BASE_COL + 3)
            .string("Totales / Suma")
            .style(asignaturaStyle);
          ws.cell(2, BASE_COL + 4)
            .string("Promedio")
            .style(asignaturaStyle);

          //RESUMEN
          ws.cell(1, BASE_COL + 3, 1, BASE_COL + 4, true)
            .string("RESUMEN DE HORAS")
            .style(header);
          //ASIGNATURAS
          ws.cell(
            1,
            BASE_COL + 5,
            1,
            BASE_COL + 4 + dataFormated[index].cursos.length,
            true
          )
            .string("ASIGNATURAS")
            .style(header);
          dataFormated[index].cursos.forEach((asignatura, i) => {
            ws.column(BASE_COL + 5 + i).setWidth(15);
            ws.cell(2, BASE_COL + 5 + i)
              .string(asignatura.title)
              .style(asignaturaStyle);
          });

          //   alumnos
          dataFormated[index].alumnos
            .filter(
              (alumno) =>
                alumno.status == 1 ||
                (alumno.status == 5 &&
                  alumno.username.search("e") != -1 &&
                  alumno.username.search("E") != -1)
            )
            .forEach((alumno, i) => {
              ws.cell(3 + i, 1)
                .number(i + 1)
                .style(i % 2 == 0 ? alumnoStyle1 : alumnoStyle2);
              ws.cell(3 + i, 2)
                .string(alumno.lastname)
                .style(i % 2 == 0 ? alumnoStyle1 : alumnoStyle2);
              ws.cell(3 + i, 3)
                .string(alumno.firstname)
                .style(i % 2 == 0 ? alumnoStyle1 : alumnoStyle2);

              let totalTime = 0;

              // TOTAL DE HORAS
              dataFormated[index].cursos.forEach((asignatura, j) => {
                let currentTime = alumno.asignaturas.find(
                  (item) => item.c_id == asignatura.id
                );

                if (!currentTime) return;

                totalTime = totalTime + timeFix(currentTime.time);
              });

              // TOTAL
              ws.cell(3 + i, 4)
                .string(secondsFix(totalTime))
                .style({
                  ...valorStyle,
                  ...getBackground(secondsFix(totalTime)),
                  numberFormat: "[h]:mm:ss",
                });
              // PROMEDIO
              ws.cell(3 + i, 5)
                .string(
                  secondsFix(parseInt(totalTime / alumno.asignaturas.length))
                )
                .style({
                  ...valorStyle,
                  ...getBackground(
                    secondsFix(parseInt(totalTime / alumno.asignaturas.length))
                  ),
                  numberFormat: "[h]:mm:ss",
                });

              ws.column(5).freeze();
              dataFormated[index].cursos.forEach((asignatura, j) => {
                const currentTime = alumno.asignaturas.find(
                  (item) => item.c_id == asignatura.id
                );

                ws.cell(3 + i, BASE_COL + 5 + j)
                  .string(currentTime ? currentTime.time : "00:00:00")
                  .style({
                    ...valorStyle,
                    ...getBackground(
                      currentTime ? currentTime.time : "00:00:00"
                    ),
                    numberFormat: "[h]:mm:ss",
                  });
              });
            });
        };

        fileStudent();
        fileDocente();

        // CREACION DE ARCHIVO EXCEL
        const fileCreate = await wb.writeToBuffer();
        fs.writeFileSync(
          `${dataFormated[index].name.replace(/ /g, "")}.xlsx`,
          fileCreate
        );
      }

      // COMPRESION DE DATOS
      var output = fs.createWriteStream(
        `./informe_${moment().format("DD_MM_YYYY_HH_mm_ss")}${SCHOOL}.zip`
      );
      var archive = archiver("zip", {
        gzip: true,
        zlib: { level: 9 }, // Sets the compression level.
      });

      archive.on("error", function (err) {
        console.log(err);
        throw err;
      });
      // pipe archive data to the output file
      archive.pipe(output);

      // AGREGAR CADA ELEMENTO
      for (let index = 0; index < dataFormated.length; index++) {
        // append files
        await archive.file(
          `${dataFormated[index].name.replace(/ /g, "")}.xlsx`,
          {
            name: `${dataFormated[index].name.replace(/ /g, "")}.xlsx`,
          }
        );
      }

      await archive.finalize();

      // ELIMINAR CADA ELEMENTO
      for (let index = 0; index < dataFormated.length; index++) {
        fs.unlinkSync(`${dataFormated[index].name.replace(/ /g, "")}.xlsx`);
      }
    } finally {
      mysqlssh.close();
    }
  })
  .catch((err) => {
    console.log(err);
  });

// RESUMEN POR CURSOS
// const resuemnTotal = async () => {
//   // Create a new instance of a Workbook class
//   let wb = new xl.Workbook({});
//   // ESTILOS
//   // Create a reusable style
//   var header = wb.createStyle({
//     border: {
//       // §18.8.4 border (Border)
//       left: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       right: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       top: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       bottom: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//     },
//     fill: {
//       type: "pattern",
//       patternType: "solid",
//       bgColor: "#0277bd",
//       fgColor: "#0277bd",
//     },
//     font: {
//       bold: true,
//       color: "#ffffff",
//       size: 10,
//     },
//     alignment: {
//       vertical: "center",
//       // horizontal: "center",
//     },
//   });
//   var asignaturaStyle = wb.createStyle({
//     border: {
//       // §18.8.4 border (Border)
//       left: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       right: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       top: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       bottom: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//     },
//     fill: {
//       type: "pattern",
//       patternType: "solid",
//       bgColor: "#D9E1F2",
//       fgColor: "#D9E1F2",
//     },
//     font: {
//       bold: true,
//       // color: "#ffffff",
//       size: 10,
//     },
//     alignment: {
//       vertical: "center",
//       horizontal: "center",
//       wrapText: true,
//     },
//   });
//   var styleTotal = wb.createStyle({
//     border: {
//       // §18.8.4 border (Border)
//       left: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       right: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       top: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       bottom: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//     },
//     fill: {
//       type: "pattern",
//       patternType: "solid",
//       bgColor: "#D9D9D9",
//       fgColor: "#D9D9D9",
//     },
//     font: {
//       bold: true,
//       // color: "#ffffff",
//       size: 10,
//     },
//     alignment: {
//       vertical: "center",
//       horizontal: "right",
//       wrapText: true,
//     },
//   });
//   var alumnoStyle1 = wb.createStyle({
//     border: {
//       // §18.8.4 border (Border)
//       left: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       right: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       top: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       bottom: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//     },
//     fill: {
//       type: "pattern",
//       patternType: "solid",
//       bgColor: "#D9D9D9",
//       fgColor: "#D9D9D9",
//     },
//     font: {
//       bold: true,
//       // color: "#ffffff",
//       size: 10,
//     },
//     alignment: {
//       vertical: "center",
//       // horizontal: "center",
//       wrapText: true,
//     },
//   });
//   var alumnoStyle2 = wb.createStyle({
//     border: {
//       // §18.8.4 border (Border)
//       left: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       right: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       top: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       bottom: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//     },

//     font: {
//       bold: true,
//       // color: "#ffffff",
//       size: 10,
//     },
//     alignment: {
//       vertical: "center",
//       // horizontal: "center",
//       wrapText: true,
//     },
//   });
//   var valorStyle = wb.createStyle({
//     border: {
//       // §18.8.4 border (Border)
//       left: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       right: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       top: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//       bottom: {
//         style: "medium", //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
//         color: "#000000", // HTML style hex value
//       },
//     },
//     // fill: {
//     //   type: "pattern",
//     //   patternType: "solid",
//     //   bgColor: "#4fc3f7",
//     //   fgColor: "#4fc3f7",
//     // },
//     font: {
//       // color: "#ffffff",
//       size: 12,
//     },
//     alignment: {
//       horizontal: "center",
//       wrapText: true,
//     },
//   });

//   const fileDocente = async (title, type) => {
//     // Add Worksheets to the workbook
//     let ws = wb.addWorksheet(title);

//     const columnas = [
//       { name: "Cursos", index: 1, width: 30 },
//       { name: " Sin conexión", index: 2, width: 25 },
//       {
//         name: "Conexión insuficiente",
//         index: 3,
//         width: 25,
//         label: " (2 segundos - 30 minutos)",
//       },
//       {
//         name: "Conexión Media",
//         index: 4,
//         width: 25,
//         label: " (30 minutos - 1 hora)",
//       },
//       {
//         name: "Conexión Adecuada",
//         index: 5,
//         width: 25,
//         label: " (Más de una hora)",
//       },
//     ];

//     // HEADERS
//     let BASE_COL = 1;
//     ws.row(1).setHeight(30);
//     ws.cell(1, BASE_COL, 1, 5, true).string(title).style(header);

//     // Columns
//     columnas.forEach((item, i) => {
//       ws.column(BASE_COL + i).setWidth(item.width);
//       // ="Conexión insuficiente" &  CARACTER(10)  &  "(2 segundos - 30 minutos)"

//       ws.cell(2, BASE_COL + i)
//         .string(`${item.name} ${item.label ? item.label : ""}`)
//         .style(asignaturaStyle);
//     });

//     //filas
//     dataFormated
//       .sort((a, b) =>
//         a.tree_pos < b.tree_pos ? -1 : a.tree_pos > b.tree_pos ? 1 : 0
//       )
//       .forEach((curso, i) => {
//         let dnt = 0;
//         let low = 0;
//         let mid = 0;
//         let high = 0;

//         ws.cell(3 + i, 1)
//           .string(curso.name)
//           .style(i % 2 == 0 ? alumnoStyle1 : alumnoStyle2);
//         curso.alumnos
//           .filter((item) =>
//             type == 5
//               ? item.status == type && item.isTeacher != 1
//               : item.status == type
//           )
//           .forEach((alumno) => {
//             let currentTime = 0;

//             alumno.asignaturas.forEach((cu) => {
//               currentTime = currentTime + timeFix(cu.time);
//             });

//             if (currentTime <= 1) dnt = dnt + 1;
//             else if (currentTime >= 2 && currentTime <= 1800)
//               low = low + 1;
//             else if (currentTime >= 1801 && currentTime <= 3599)
//               mid = mid + 1;
//             else high = high + 1;
//           });

//         ws.cell(3 + i, 2)
//           .number(dnt)
//           .style(i % 2 == 0 ? alumnoStyle1 : alumnoStyle2);
//         ws.cell(3 + i, 3)
//           .number(low)
//           .style(i % 2 == 0 ? alumnoStyle1 : alumnoStyle2);
//         ws.cell(3 + i, 4)
//           .number(mid)
//           .style(i % 2 == 0 ? alumnoStyle1 : alumnoStyle2);
//         ws.cell(3 + i, 5)
//           .number(high)
//           .style(i % 2 == 0 ? alumnoStyle1 : alumnoStyle2);
//       });

//     // TOTAL
//     ws.cell(2 + dataFormated.length + 2, 1)
//       .string(`TOTALES`)
//       .style(styleTotal);
//     columnas.forEach((item, i) => {
//       if (i > 0) {
//         ws.cell(2 + dataFormated.length + 2, BASE_COL + i)
//           .formula(
//             `SUM(${xl.getExcelCellRef(
//               3,
//               BASE_COL + i
//             )}:${xl.getExcelCellRef(
//               2 + dataFormated.length,
//               BASE_COL + i
//             )})`
//           )
//           .style(asignaturaStyle);
//       }
//     });
//     // PORCENTAJES
//     ws.cell(2 + dataFormated.length + 3, 1)
//       .string(`PORCENTAJES`)
//       .style(styleTotal);
//     columnas.forEach((item, i) => {
//       if (i > 0) {
//         ws.cell(2 + dataFormated.length + 3, BASE_COL + i)
//           .formula(
//             `(${xl.getExcelCellRef(
//               2 + dataFormated.length + 2,
//               BASE_COL + i
//             )})/SUM(${xl.getExcelCellRef(
//               2 + dataFormated.length + 2,
//               2
//             )}:${xl.getExcelCellRef(
//               2 + dataFormated.length + 2,
//               columnas.length
//             )})`
//           )
//           .style({
//             numberFormat: "0%",
//             ...asignaturaStyle,
//             fill: {
//               type: "pattern",
//               patternType: "solid",
//               bgColor: "#D9E1F2",
//               fgColor: "#D9E1F2",
//             },
//           });
//       }
//     });
//   };
//   fileDocente("Docentes", 1);
//   fileDocente("Alumnos", 5);

//   // CREACION DE ARCHIVO EXCEL
//   const fileCreate = await wb.writeToBuffer();
//   fs.writeFileSync(`Resumen_acceso.xlsx`, fileCreate);
// };
